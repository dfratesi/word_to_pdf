# version de Claude

import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
import os
from pathlib import Path
import threading
import queue
import traceback
import tempfile
import time

from concurrent.futures import ThreadPoolExecutor, TimeoutError
import subprocess


# --- Dependencias específicas de Windows ---
try:
    from docx2pdf import convert as convert_windows_docx2pdf
    import win32com.client
    import pythoncom

    WINDOWS_DEPS_OK = True
except ImportError:
    WINDOWS_DEPS_OK = False


def log_message(text_widget, message):
    if text_widget:
        try:
            text_widget.configure(state=tk.NORMAL)
            text_widget.insert(tk.END, str(message) + "\n")
            text_widget.see(tk.END)
            text_widget.configure(state=tk.DISABLED)
            text_widget.update_idletasks()
        except Exception as e:
            print(f"Error al loguear en GUI: {e}")


def convert_doc_to_docx_with_timeout(doc_path, docx_path, timeout_seconds=60):
    """
    Convierte un archivo .doc a .docx con timeout
    """

    def convert_doc_to_docx():
        try:
            # Inicializar COM para este hilo
            pythoncom.CoInitialize()

            # Crear instancia de Word
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = False
            word_app.DisplayAlerts = 0  # No mostrar alertas

            # Configurar Word para que no muestre diálogos
            word_app.Options.DoNotPromptForConvert = True
            word_app.Options.ConfirmConversions = False

            try:
                # Abrir documento
                doc = word_app.Documents.Open(
                    doc_path, ConfirmConversions=False, ReadOnly=True
                )

                # Guardar como .docx
                doc.SaveAs2(
                    docx_path, FileFormat=12
                )  # 12 = wdFormatXMLDocument (.docx)
                doc.Close(SaveChanges=0)  # 0 = wdDoNotSaveChanges

                return True

            except Exception as e:
                raise Exception(f"Error durante conversión: {e}")
            finally:
                try:
                    word_app.Quit()
                except:
                    pass

        except Exception as e:
            raise Exception(f"Error en proceso de conversión: {e}")
        finally:
            try:
                pythoncom.CoUninitialize()
            except:
                pass

    # Ejecutar con timeout
    with ThreadPoolExecutor(max_workers=1) as executor:
        future = executor.submit(convert_doc_to_docx)
        try:
            result = future.result(timeout=timeout_seconds)
            return True
        except TimeoutError:
            raise Exception(
                f"Timeout de {timeout_seconds} segundos alcanzado durante conversión .doc a .docx"
            )
        except Exception as e:
            raise e


def kill_word_processes():
    """
    Mata todos los procesos de Word que puedan estar colgados
    """
    try:
        subprocess.run(
            ["taskkill", "/f", "/im", "WINWORD.EXE"],
            capture_output=True,
            text=True,
            timeout=10,
        )
    except:
        pass


def convertir_archivos_worker_windows(
    ruta_carpeta_entrada, ruta_carpeta_salida, text_widget_log, progress_queue
):
    archivos_convertidos = 0
    archivos_fallidos = 0
    total_archivos_word = 0

    log_message(
        text_widget_log,
        f"Entrada worker: Carpeta Word='{ruta_carpeta_entrada}', Carpeta PDF='{ruta_carpeta_salida}'",
    )

    try:
        ruta_carpeta_entrada_abs = str(Path(ruta_carpeta_entrada).resolve())
        ruta_carpeta_salida_abs = str(Path(ruta_carpeta_salida).resolve())

        log_message(
            text_widget_log,
            f"Rutas absolutas worker: Entrada='{ruta_carpeta_entrada_abs}', Salida='{ruta_carpeta_salida_abs}'",
        )
        lista_archivos = os.listdir(ruta_carpeta_entrada_abs)
    except Exception as e:
        log_message(
            text_widget_log,
            f"ERROR: No se pudo leer la carpeta de entrada '{ruta_carpeta_entrada_abs}': {e}",
        )
        log_message(text_widget_log, traceback.format_exc())
        progress_queue.put(
            {"type": "finish_error", "message": "Fallo lectura carpeta entrada"}
        )
        return

    # Contar archivos válidos
    for nombre_archivo in lista_archivos:
        if nombre_archivo.lower().endswith(
            (".docx", ".doc")
        ) and not nombre_archivo.startswith("~$"):
            total_archivos_word += 1

    if total_archivos_word == 0:
        log_message(
            text_widget_log,
            "No se encontraron archivos .doc o .docx válidos en la carpeta de entrada.",
        )
        progress_queue.put({"type": "finish"})
        return

    archivos_procesados = 0

    # Verificación inicial de que Word COM funciona
    try:
        log_message(text_widget_log, "Verificación rápida de MS Word COM...")
        pythoncom.CoInitialize()
        test_word_app = win32com.client.Dispatch("Word.Application")
        test_word_app.Visible = False
        test_word_app.Quit()
        pythoncom.CoUninitialize()
        log_message(text_widget_log, "MS Word COM respondió a la verificación inicial.")
    except Exception as e:
        log_message(
            text_widget_log,
            f"ERROR CRÍTICO: Fallo en la verificación inicial de MS Word COM.",
        )
        log_message(text_widget_log, f"Detalles: {e}\n{traceback.format_exc()}")
        progress_queue.put(
            {"type": "finish_error", "message": "Fallo MS Word COM inicial"}
        )
        return

    # Procesar cada archivo
    for nombre_archivo in lista_archivos:
        if nombre_archivo.startswith("~$"):
            continue

        ruta_completa_archivo_original = os.path.join(
            ruta_carpeta_entrada_abs, nombre_archivo
        )
        path_original_resolved = str(Path(ruta_completa_archivo_original).resolve())

        if os.path.isfile(path_original_resolved) and (
            nombre_archivo.lower().endswith(".docx")
            or nombre_archivo.lower().endswith(".doc")
        ):
            archivos_procesados += 1
            nombre_base_archivo = Path(nombre_archivo).stem
            ruta_archivo_pdf_resolved = str(
                Path(ruta_carpeta_salida_abs) / f"{nombre_base_archivo}.pdf"
            )

            log_message(
                text_widget_log,
                f"Procesando '{nombre_archivo}' ({archivos_procesados}/{total_archivos_word})...",
            )
            log_message(text_widget_log, f"  Ruta original: '{path_original_resolved}'")

            path_to_convert_for_docx2pdf = path_original_resolved
            temp_docx_file_path = None
            conversion_successful = False

            # Manejar archivos .doc
            if nombre_archivo.lower().endswith(".doc"):
                log_message(
                    text_widget_log,
                    f"  '{nombre_archivo}' es .doc. Iniciando pre-conversión a .docx temporal (timeout: 60s)...",
                )

                try:
                    # Crear archivo temporal
                    fd, temp_docx_file_path_raw = tempfile.mkstemp(
                        suffix=".docx", prefix=f"{nombre_base_archivo}_temp_"
                    )
                    os.close(fd)
                    temp_docx_file_path = str(Path(temp_docx_file_path_raw).resolve())

                    log_message(
                        text_widget_log,
                        f"    Convirtiendo a: '{temp_docx_file_path}'",
                    )

                    # Intentar conversión con timeout
                    start_time = time.time()
                    convert_doc_to_docx_with_timeout(
                        path_original_resolved, temp_docx_file_path, timeout_seconds=60
                    )
                    end_time = time.time()

                    log_message(
                        text_widget_log,
                        f"  Pre-conversión completada en {end_time - start_time:.1f} segundos",
                    )
                    path_to_convert_for_docx2pdf = temp_docx_file_path

                except Exception as pre_conv_e:
                    log_message(
                        text_widget_log,
                        f"  ERROR durante pre-conversión .doc a .docx para '{nombre_archivo}': {pre_conv_e}",
                    )

                    # Intentar matar procesos de Word colgados
                    log_message(
                        text_widget_log,
                        "  Intentando limpiar procesos de Word colgados...",
                    )
                    kill_word_processes()
                    time.sleep(2)

                    archivos_fallidos += 1

                    # Limpiar archivo temporal si existe
                    if temp_docx_file_path and os.path.exists(temp_docx_file_path):
                        try:
                            os.remove(temp_docx_file_path)
                        except:
                            pass

                    # Actualizar progreso y continuar con siguiente archivo
                    progress_queue.put(
                        {
                            "type": "progress",
                            "value": archivos_procesados / total_archivos_word,
                        }
                    )
                    continue

            # Convertir a PDF usando docx2pdf
            try:
                log_message(
                    text_widget_log,
                    f"  Convirtiendo a PDF: '{Path(path_to_convert_for_docx2pdf).name}' -> '{Path(ruta_archivo_pdf_resolved).name}'",
                )

                start_time = time.time()
                convert_windows_docx2pdf(
                    path_to_convert_for_docx2pdf, ruta_archivo_pdf_resolved
                )
                end_time = time.time()

                if (
                    os.path.exists(ruta_archivo_pdf_resolved)
                    and os.path.getsize(ruta_archivo_pdf_resolved) > 0
                ):
                    log_message(
                        text_widget_log,
                        f"  ✓ CONVERTIDO exitosamente en {end_time - start_time:.1f} segundos",
                    )
                    archivos_convertidos += 1
                    conversion_successful = True
                else:
                    log_message(
                        text_widget_log,
                        "  ✗ ERROR: PDF no generado o está vacío",
                    )
                    archivos_fallidos += 1

            except Exception as e:
                log_message(
                    text_widget_log,
                    f"  ✗ ERROR al convertir a PDF: {e}",
                )
                archivos_fallidos += 1

            finally:
                # Limpiar archivo temporal
                if temp_docx_file_path and os.path.exists(temp_docx_file_path):
                    try:
                        os.remove(temp_docx_file_path)
                        log_message(
                            text_widget_log,
                            "  Archivo temporal eliminado",
                        )
                    except Exception as del_e:
                        log_message(
                            text_widget_log,
                            f"  Advertencia: No se pudo eliminar archivo temporal: {del_e}",
                        )

            # Actualizar progreso
            progress_queue.put(
                {"type": "progress", "value": archivos_procesados / total_archivos_word}
            )

    # Resumen final
    log_message(text_widget_log, "\n" + "=" * 50)
    log_message(text_widget_log, "RESUMEN DE CONVERSIÓN")
    log_message(text_widget_log, "=" * 50)
    log_message(text_widget_log, f"Archivos Word encontrados: {total_archivos_word}")
    log_message(text_widget_log, f"Archivos procesados: {archivos_procesados}")
    log_message(text_widget_log, f"Conversiones exitosas: {archivos_convertidos}")
    log_message(text_widget_log, f"Archivos con errores: {archivos_fallidos}")

    if archivos_fallidos > 0:
        log_message(
            text_widget_log,
            f"\n⚠️  {archivos_fallidos} archivos fallaron en la conversión",
        )
        log_message(text_widget_log, "Posibles causas:")
        log_message(text_widget_log, "- Archivos corruptos o protegidos")
        log_message(text_widget_log, "- Contenido complejo que Word no puede procesar")
        log_message(text_widget_log, "- Problemas de permisos")

    progress_queue.put({"type": "finish"})


class WordToPdfConverterApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Convertidor Word a PDF (Windows) - Versión Mejorada")
        self.geometry("800x650")
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.ruta_carpeta_entrada = tk.StringVar()
        self.ruta_carpeta_salida = tk.StringVar()
        self.progress_queue = queue.Queue()

        # Frame para selección de carpetas
        self.frame_carpetas = ctk.CTkFrame(self)
        self.frame_carpetas.pack(pady=10, padx=20, fill="x")

        ctk.CTkLabel(self.frame_carpetas, text="Carpeta Word (Entrada):").grid(
            row=0, column=0, padx=5, pady=5, sticky="w"
        )
        self.entry_entrada = ctk.CTkEntry(
            self.frame_carpetas, textvariable=self.ruta_carpeta_entrada, width=400
        )
        self.entry_entrada.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(
            self.frame_carpetas,
            text="Seleccionar",
            command=self.seleccionar_carpeta_entrada,
        ).grid(row=0, column=2, padx=5, pady=5)

        ctk.CTkLabel(self.frame_carpetas, text="Carpeta PDF (Salida):").grid(
            row=1, column=0, padx=5, pady=5, sticky="w"
        )
        self.entry_salida = ctk.CTkEntry(
            self.frame_carpetas, textvariable=self.ruta_carpeta_salida, width=400
        )
        self.entry_salida.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(
            self.frame_carpetas,
            text="Seleccionar",
            command=self.seleccionar_carpeta_salida,
        ).grid(row=1, column=2, padx=5, pady=5)

        self.frame_carpetas.grid_columnconfigure(1, weight=1)

        # Botones de control
        self.frame_botones = ctk.CTkFrame(self)
        self.frame_botones.pack(pady=10)

        self.convert_button = ctk.CTkButton(
            self.frame_botones,
            text="Iniciar Conversión",
            command=self.iniciar_conversion,
            height=35,
            width=150,
        )
        self.convert_button.pack(side="left", padx=5)

        self.clear_log_button = ctk.CTkButton(
            self.frame_botones,
            text="Limpiar Log",
            command=self.limpiar_log,
            height=35,
            width=100,
        )
        self.clear_log_button.pack(side="left", padx=5)

        # Barra de progreso
        self.progressbar = ctk.CTkProgressBar(self, width=500)
        self.progressbar.set(0)
        self.progressbar.pack(pady=(0, 10))

        # Área de log
        self.log_textbox = ctk.CTkTextbox(
            self, width=750, height=350, state=tk.DISABLED, wrap=tk.WORD
        )
        self.log_textbox.pack(pady=10, padx=20, fill="both", expand=True)

        self.verificar_dependencias_windows()
        self.after(100, self.procesar_cola_progreso)

    def verificar_dependencias_windows(self):
        log_message(self.log_textbox, "🔍 Verificando dependencias para Windows...")
        if WINDOWS_DEPS_OK:
            log_message(
                self.log_textbox,
                "✓ Plataforma: Windows. Método: Microsoft Word + docx2pdf",
            )
            log_message(
                self.log_textbox, "✓ Dependencias 'docx2pdf' y 'pywin32' encontradas"
            )
            log_message(
                self.log_textbox, "ℹ️  Requiere: Microsoft Word instalado y activado"
            )
            log_message(self.log_textbox, "ℹ️  Timeout por archivo: 60 segundos")
        else:
            log_message(
                self.log_textbox,
                "❌ ERROR CRÍTICO: Faltan bibliotecas 'docx2pdf' y/o 'pywin32'",
            )
            log_message(
                self.log_textbox,
                "📝 Instálalas ejecutando: pip install docx2pdf pywin32",
            )
            self.convert_button.configure(state=tk.DISABLED)
            messagebox.showerror(
                "Error de Dependencias",
                "Faltan 'docx2pdf' o 'pywin32'.\nInstálalas con 'pip install docx2pdf pywin32' y reinicia.",
            )
        log_message(self.log_textbox, "-" * 60)

    def seleccionar_carpeta_entrada(self):
        ruta = filedialog.askdirectory(title="Seleccionar Carpeta con Archivos Word")
        if ruta:
            ruta_normalizada = str(Path(ruta).resolve())
            self.ruta_carpeta_entrada.set(ruta_normalizada)
            if not self.ruta_carpeta_salida.get():
                self.ruta_carpeta_salida.set(
                    str(Path(ruta_normalizada) / "PDF_Convertidos")
                )
            log_message(self.log_textbox, f"📁 Carpeta de entrada: {ruta_normalizada}")

    def seleccionar_carpeta_salida(self):
        ruta = filedialog.askdirectory(title="Seleccionar Carpeta para Guardar los PDF")
        if ruta:
            ruta_normalizada = str(Path(ruta).resolve())
            self.ruta_carpeta_salida.set(ruta_normalizada)
            log_message(self.log_textbox, f"📁 Carpeta de salida: {ruta_normalizada}")

    def limpiar_log(self):
        self.log_textbox.configure(state=tk.NORMAL)
        self.log_textbox.delete(1.0, tk.END)
        self.log_textbox.configure(state=tk.DISABLED)

    def iniciar_conversion(self):
        if not WINDOWS_DEPS_OK:
            messagebox.showerror(
                "Error de Dependencias",
                "Faltan dependencias críticas.\nInstala 'docx2pdf' y 'pywin32' y reinicia.",
            )
            return

        entrada = self.ruta_carpeta_entrada.get()
        salida = self.ruta_carpeta_salida.get()

        if not entrada or not os.path.isdir(entrada):
            messagebox.showerror("Error", "Selecciona una carpeta de entrada válida.")
            return
        if not salida:
            messagebox.showerror("Error", "Selecciona una carpeta de salida.")
            return

        if not os.path.exists(salida):
            try:
                Path(salida).mkdir(parents=True, exist_ok=True)
                log_message(self.log_textbox, f"📁 Carpeta de salida creada: {salida}")
            except OSError as e:
                messagebox.showerror(
                    "Error", f"No se pudo crear la carpeta de salida: {e}"
                )
                return

        self.convert_button.configure(state=tk.DISABLED)
        self.progressbar.set(0)
        log_message(self.log_textbox, f"\n🚀 INICIANDO CONVERSIÓN")
        log_message(self.log_textbox, f"📂 Desde: {entrada}")
        log_message(self.log_textbox, f"📂 Hacia: {salida}")
        log_message(self.log_textbox, "=" * 60)

        self.conversion_thread = threading.Thread(
            target=convertir_archivos_worker_windows,
            args=(entrada, salida, self.log_textbox, self.progress_queue),
            daemon=True,
        )
        self.conversion_thread.start()

    def procesar_cola_progreso(self):
        try:
            while True:
                mensaje = self.progress_queue.get_nowait()
                if mensaje["type"] == "progress":
                    self.progressbar.set(mensaje["value"])
                elif mensaje["type"] == "finish":
                    self.progressbar.set(1)
                    self.convert_button.configure(state=tk.NORMAL)
                    messagebox.showinfo(
                        "Completado",
                        "✅ Conversión finalizada.\nRevisa el log para detalles.",
                    )
                    break
                elif mensaje["type"] == "finish_error":
                    self.progressbar.set(0)
                    self.convert_button.configure(state=tk.NORMAL)
                    messagebox.showerror(
                        "Error en Conversión",
                        f"❌ Error crítico: {mensaje.get('message', 'Desconocido')}\nRevisa el log.",
                    )
                    break
        except queue.Empty:
            pass
        finally:
            self.after(100, self.procesar_cola_progreso)


if __name__ == "__main__":
    app = WordToPdfConverterApp()
    app.mainloop()
