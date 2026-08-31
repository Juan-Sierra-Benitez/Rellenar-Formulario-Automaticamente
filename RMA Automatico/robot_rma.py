import os
import sys
import time
import json
import tkinter as tk
from tkinter import messagebox, filedialog
import pyautogui
import win32com.client
import re
import pyperclip
from datetime import datetime

CONFIG_FILE = "coordenadas_bot.json"

def get_outlook_data():
    try:
        outlook_app = win32com.client.Dispatch('Outlook.Application')
        explorer = outlook_app.ActiveExplorer()
        
        if explorer is None:
            if outlook_app.Explorers.Count > 0:
                explorer = outlook_app.Explorers.Item(1)
            else:
                return {}, "No se encontró la ventana principal de Outlook."
                
        selection = explorer.Selection
        if selection.Count == 0:
            return {}, "No hay ningún correo seleccionado en Outlook. Haz clic en uno."
            
        mail = selection.Item(1)
        cuerpo = mail.Body
        asunto = mail.Subject
        
        modelo = ""
        sn = ""
        averia = ""
        rma = ""
        
        match_rma = re.search(r'RMA\s*(\d+)', asunto, re.IGNORECASE)
        if match_rma:
            rma = match_rma.group(1)
            
        # 1. Extraer Modelo
        m_mod1 = re.search(r'\b((?:HA|CS|KW)-[A-Z0-9\-]+)\b', cuerpo, re.IGNORECASE)
        if m_mod1:
            modelo = m_mod1.group(1)
        else:
            m_mod2 = re.search(r'\b(?:MODELO|MOD\.?|ARTICULO|ART\.?)\s*(?:ES\s*:?|:)?\s*([A-Z0-9\-]{4,20})\b', cuerpo, re.IGNORECASE)
            if m_mod2:
                modelo = m_mod2.group(1)
                
        # 2. Extraer Numero de Serie
        m_sn = re.search(r'\b(?:N/S|NS|N\.?S\.?|N[Oº]\s*DE\s*SERIE|N[UÚ]MERO\s*DE\s*SERIE|S/N)\s*(?:ES\s*:?|:|-)?\s*([A-Z0-9\-]{5,20})\b', cuerpo, re.IGNORECASE)
        if m_sn:
            sn = m_sn.group(1)
            
        # 3. Extraer Avería
        # Para evitar que capture "Avería: Hacer un paquete..." en el aviso legal del fondo del correo, 
        # leemos el correo de ARRIBA a ABAJO. Primera oración que encontremos, la guardamos y cortamos.
        frases_averia = []
        frases = re.split(r'[.\n]', cuerpo) # Separar por puntos o saltos de línea
        
        for f in frases:
            f = " ".join(f.split())
            if not f: continue
            f_up = f.upper()
            
            # Cortafuegos: si llegamos a la zona del aviso legal o firmas, paramos de leer.
            if "CONFIDENCIAL" in f_up or "AVISO LEGAL" in f_up or "ESTE MENSAJE" in f_up or "SALUDO" in f_up or "ATENTAMENTE" in f_up:
                break
                
            # Si alguien pone explícitamente "Avería: XXX", lo pillamos directo.
            m_strict = re.search(r'(?:AVER[ÍI]A|FALLO|S[ÍI]NTOMA|DEFECTO|PROBLEMA)\s*:\s*(.+)', f, re.IGNORECASE)
            if m_strict:
                frases_averia.append(m_strict.group(1).strip())
                break
                
            # Si no, miramos si la frase contiene las palabras malditas del cliente ("no carga", "no conecta", etc)
            if re.search(r'\b(no\s+carga|no\s+funciona|no\s+enciende|roto|falla|fallo|no\s+conecta|no\s+se\s+escucha|no\s+se\s+oye|mensaje\b|error\b)\b', f, re.IGNORECASE):
                if len(f) > 5 and len(f) < 250:
                    frases_averia.append(f)
                    # No hacemos break para poder pillar otra frase junta si la hubiera (y luego juntarlas)
                    
        if frases_averia:
            averia = " ".join(frases_averia)
                        
        # Limpieza final
        modelo = modelo.replace(":", "").replace(",", "").replace(".", "").strip().upper()
        sn = sn.replace(":", "").replace(",", "").replace(".", "").strip().upper()
        averia = averia.strip().upper()
        rma = rma.strip().upper()

        return {"modelo": modelo, "sn": sn, "averia": averia, "rma": rma}, "OK"
    except Exception as e:
        return {}, f"Error al leer Outlook: {str(e)}\nAsegúrate de tener Outlook abierto."

class Calibrador:
    def __init__(self, master):
        self.master = master
        self.puntos = [
            ("Fecha recepción", "Haz clic en el cuadro de 'Fecha recepción'"),
            ("Día de hoy (Calendario)", "¡OJO! Usa tu ratón para CLICAR NORMALMENTE el cuadro anterior para mostrar el calendario, quita el error de Albarán, y luego pon el ratón sobre EL DÍA DE HOY (Ej: 18) y guarda con CTRL."),
            ("Técnico Borrador", "Haz clic en el cuadro de 'Técnico Borrador'"),
            ("Recepcionista", "Haz clic en el cuadro de 'Recepcionista'"),
            ("Cliente", "Haz clic en el cuadro de 'Cliente'"),
            ("Marca", "Haz clic en el cuadro de 'Marca'"),
            ("Modelo", "Haz clic en el cuadro de 'Modelo'"),
            ("Número de serie", "Haz clic en el cuadro de 'Número de serie'"),
            ("Tipo de servicio", "Haz clic en el recuadro amarillo de 'Tipo de servicio'"),
            ("Garantía", "Haz clic en el cuadro de '¿Está en garantía?' (el botón)"),
            ("Fecha de compra", "Haz clic en el cuadro de 'Fecha de compra'"),
            ("Su referencia", "Haz clic en el cuadro de 'Su referencia'"),
            ("Botón ?", "Haz clic en el BOTÓN AZUL '?' de Accesorios"),
            ("Botón Usado (Popup)", "¡OJO! Haz clic AHORA MISMO en el botón '?' para abrir la ventana.\nLuego pon el ratón y guarda el botón de 'USADO'."),
            ("Obs Generales (Popup)", "Haz clic en el cuadro de 'Observaciones generales' del popup."),
            ("Guardar (Popup)", "Haz clic en el botón 'Guardar' del popup."),
            ("Observaciones usuario", "Haz clic en el cuadro final de 'Observaciones usuario' (la avería)")
        ]
        self.current_idx = 0
        self.coordenadas = {}
        self.listener = None
        
        self.top = tk.Toplevel(master)
        self.top.attributes("-topmost", True)
        self.top.geometry("750x150+50+50")
        self.top.title("Calibrador del Robot")
        
        self.lbl = tk.Label(self.top, text="", font=("Segoe UI", 12, "bold"), fg="red", wraplength=700)
        self.lbl.pack(expand=True, fill="both", padx=10, pady=10)
        
    def start(self):
        messagebox.showinfo("INFO", "Vamos a calibrar los clics del Robot.\nPor cada paso, mueve el cursor SIN HACER CLIC sobre el lugar indicado.\nLuego, pulsa la tecla CONTROL (Ctrl) de tu teclado para guardar esa posición.", parent=self.top)
        self.next_point()
        
    def next_point(self):
        if self.current_idx < len(self.puntos):
            nombre, msg = self.puntos[self.current_idx]
            self.lbl.config(text=f"Paso {self.current_idx+1}/{len(self.puntos)}: {msg}\n(Pon el cursor encima y pulsa CTRL para guardar)")
            
            from pynput.keyboard import Listener, Key
            def on_press(key):
                if key == Key.ctrl_l or key == Key.ctrl_r:
                    x, y = pyautogui.position()
                    self.coordenadas[nombre] = (int(x), int(y))
                    return False
            
            self.listener = Listener(on_press=on_press)
            self.listener.start()
            self.master.after(100, self.check_listener)
        else:
            self.coordenadas["last_calibrated_date"] = datetime.now().strftime("%d/%m/%Y")
            with open(CONFIG_FILE, 'w') as f:
                json.dump(self.coordenadas, f)
            self.top.destroy()
            messagebox.showinfo("Listo", "Calibración terminada.\nEl Robot ya sabe dónde hacer los clics.")

    def check_listener(self):
        if self.listener and not self.listener.is_alive():
            self.current_idx += 1
            self.master.after(600, self.next_point)
        else:
            self.master.after(100, self.check_listener)


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("🤖 Asistente RMA Sotelec")
        self.root.geometry("450x650")
        
        f_out = tk.LabelFrame(root, text=" 📩 Datos desde Outlook ", font=("Arial", 10, "bold"))
        f_out.pack(pady=10, padx=15, fill="x")
        
        btn_leer = tk.Button(f_out, text="🔄 1. LEER CORREO SELECCIONADO", bg="#ADD8E6", font=("Arial", 10, "bold"), command=self.leer_correo)
        btn_leer.pack(pady=10, fill="x", padx=10)
        
        self.lbl_status = tk.Label(f_out, text="Sin leer...", fg="grey")
        self.lbl_status.pack()
        
        def crear_fila(padre, texto):
            f = tk.Frame(padre)
            f.pack(fill="x", padx=10, pady=2)
            tk.Label(f, text=texto, width=12, anchor="w").pack(side="left")
            var = tk.StringVar()
            # SE HAN PUESTO EN NORMAL PARA PODER EDITARLOS MANUALMENTE SI HACE FALTA
            tk.Entry(f, textvariable=var, state="normal").pack(side="left", fill="x", expand=True)
            return var
            
        self.var_modelo = crear_fila(f_out, "Modelo:")
        self.var_sn = crear_fila(f_out, "Num Serie:")
        self.var_averia = crear_fila(f_out, "Avería:")
        
        f = tk.Frame(f_out)
        f.pack(fill="x", padx=10, pady=2)
        tk.Label(f, text="Nº RMA:", width=12, anchor="w").pack(side="left")
        self.var_rma = tk.StringVar()
        tk.Entry(f, textvariable=self.var_rma).pack(side="left", fill="x", expand=True)
        
        f_man = tk.LabelFrame(root, text=" ✍️ Datos Manuales ", font=("Arial", 10, "bold"))
        f_man.pack(pady=10, padx=15, fill="x")
        
        f1 = tk.Frame(f_man)
        f1.pack(fill="x", padx=10, pady=5)
        tk.Label(f1, text="Fecha compra:", width=12, anchor="w").pack(side="left")
        self.var_fecha = tk.StringVar()
        tk.Entry(f1, textvariable=self.var_fecha).pack(side="left", fill="x", expand=True)
        
        f2 = tk.Frame(f_man)
        f2.pack(fill="x", padx=10, pady=5)
        tk.Label(f2, text="¿De dónde es?\n(Ej: ECI CALLAO):", width=12, anchor="w").pack(side="left")
        self.var_donde = tk.StringVar()
        tk.Entry(f2, textvariable=self.var_donde).pack(side="left", fill="x", expand=True)
        
        tk.Label(f_man, text="Accesorios incluidos:", anchor="w").pack(anchor="w", padx=10, pady=(10,0))
        f_acc = tk.Frame(f_man)
        f_acc.pack(fill="x", padx=10, pady=5)
        self.var_cable = tk.BooleanVar()
        self.var_manual = tk.BooleanVar()
        self.var_caja = tk.BooleanVar()
        tk.Checkbutton(f_acc, text="Cable", variable=self.var_cable).pack(side="left", padx=5)
        tk.Checkbutton(f_acc, text="Manual", variable=self.var_manual).pack(side="left", padx=5)
        tk.Checkbutton(f_acc, text="Caja", variable=self.var_caja).pack(side="left", padx=5)
        
        btn_run = tk.Button(root, text="🚀 2. ¡RELLENAR EDGE!", bg="#90EE90", font=("Arial", 14, "bold"), height=2, command=self.ejecutar_robot)
        btn_run.pack(pady=15, fill="x", padx=20)
        
        btn_clear = tk.Button(root, text="🧹 LIMPIAR TODO (Para pedido nuevo)", bg="#FFE4E1", command=self.limpiar_datos)
        btn_clear.pack(pady=5, fill="x", padx=20)
        
        btn_cal_dia = tk.Button(root, text="📅 Actualizar Solo Día de Hoy (Usar cada mañana)", bg="#FFFFE0", command=self.recalibrar_solo_dia)
        btn_cal_dia.pack(pady=2, fill="x", padx=20)
        
        btn_cal = tk.Button(root, text="⚙️ CALIBRAR TODO", bg="#FFDAB9", command=self.iniciar_calibracion)
        btn_cal.pack(pady=5, fill="x", padx=20)
        
    def leer_correo(self):
        data, status = get_outlook_data()
        if status == "OK":
            self.lbl_status.config(text="¡Leído correctamente!", fg="green")
            self.var_modelo.set(data["modelo"])
            self.var_sn.set(data["sn"])
            self.var_averia.set(data["averia"])
            self.var_rma.set(data["rma"])
            
            # EL AVISO DEL GUION SI SOLO HAY UNO
            if data["modelo"].count('-') == 1:
                messagebox.showwarning("Atención con el Modelo extraído", "El modelo extraído solo tiene 1 guión (Ej: HA-EC25T).\nEn tu Edge has visto que a veces requieres añadirle letras al final para que lo detecte bien.\n\nPor favor, completa/edita el modelo TÚ MISMO a mano en la caja de texto antes de pulsar en Rellenar Edge.")
        else:
            self.lbl_status.config(text="Error al leer", fg="red")
            messagebox.showerror("Error", status)
            
    def limpiar_datos(self):
        self.var_modelo.set("")
        self.var_sn.set("")
        self.var_averia.set("")
        self.var_rma.set("")
        self.var_fecha.set("")
        self.var_donde.set("")
        self.var_cable.set(False)
        self.var_manual.set(False)
        self.var_caja.set(False)
        self.lbl_status.config(text="✨ Limpio y listo para correo nuevo", fg="black")
        
    def iniciar_calibracion(self):
        c = Calibrador(self.root)
        c.start()
        
    def recalibrar_solo_dia(self):
        messagebox.showinfo("INFO", "Saca el Calendario en Edge, quita la alarma de Albarán, y luego PON EL RATÓN SOBRE EL DÍA (Ej: 18) y aprieta la tecla CONTROL (Ctrl).")
        from pynput.keyboard import Listener, Key
        def on_press(key):
            if key == Key.ctrl_l or key == Key.ctrl_r:
                x, y = pyautogui.position()
                
                if os.path.exists(CONFIG_FILE):
                    try:
                        with open(CONFIG_FILE, 'r') as f:
                            c = json.load(f)
                    except:
                        c = {}
                else:
                    c = {}
                    
                c["Día de hoy (Calendario)"] = (int(x), int(y))
                c["last_calibrated_date"] = datetime.now().strftime("%d/%m/%Y")
                with open(CONFIG_FILE, 'w') as f:
                    json.dump(c, f)
                    
                messagebox.showinfo("¡Listo!", "Posición del día actualizada. El Robot clicará en este cuadradito hoy.")
                return False
        
        self.listener = Listener(on_press=on_press)
        self.listener.start()
        
    def ejecutar_robot(self):
        if not os.path.exists(CONFIG_FILE):
            messagebox.showwarning("Falta Calibrar", "Debes hacer la calibración primero.")
            return
            
        with open(CONFIG_FILE, 'r') as f:
            c = json.load(f)
            
        if "Día de hoy (Calendario)" not in c:
            messagebox.showerror("Recalibrar obligatorio", "¡Actualización! Ahora el robot va a Clicar físicamente en el día del calendario. Tienes que volver a darle a CALIBRAR COORDENADAS una vez más.")
            return
            
        hoy_str = datetime.now().strftime("%d/%m/%Y")
        if c.get("last_calibrated_date") != hoy_str:
            ans = messagebox.askyesno("⚠️ Aviso de Nuevo Día", f"El robot se ha dado cuenta de que hoy es {hoy_str}, pero tu recuadro del calendario se guardó un día distinto.\n\nSi dejas que el robot siga, hará clic en la casilla de ayer.\n\n¿Quieres FRENAR AL ROBOT ahora mismo para poder darle tú al NUEVO BOTÓN AMARILLO y actualizar el día?\n\n(Pulsa SÍ para parar y actualizar. Pulsa NO si por lo que sea quieres clicar en la casilla de ayer).")
            if ans:
                return
            
        # TODO A MAYUSCULAS COMO PIDE EL USUARIO
        modelo = self.var_modelo.get().upper()
        sn = self.var_sn.get().upper()
        averia = self.var_averia.get().upper()
        rma = self.var_rma.get().upper()
        fecha_compra = self.var_fecha.get().upper()
        de_donde = self.var_donde.get().upper()
        
        acc = []
        if self.var_cable.get(): acc.append("SI TIENE CABLE")
        else: acc.append("NO TIENE CABLE")
        
        if self.var_manual.get(): acc.append("SI TIENE MANUAL")
        else: acc.append("NO TIENE MANUAL")
        
        if self.var_caja.get(): acc.append("SI TIENE CAJA")
        else: acc.append("NO TIENE CAJA")
        
        obs_accesorios = ", ".join(acc).upper()
            
        hoy = datetime.now().strftime("%d/%m/%Y")
        su_ref = f"RMA {rma} - {de_donde}".upper()
        
        self.root.iconify()
        messagebox.showinfo("¡Preparado!", "Asegúrate de que Edge está visible.\n\nEl robot tomará el control. ¡NO TOQUES NADA hasta que termine!")
        time.sleep(1)
        pyautogui.FAILSAFE = True
        
        try:
            def clica(key):
                if key in c:
                    # Movimiento suave hacia el punto ayuda a que procese el hover
                    pyautogui.moveTo(c[key][0], c[key][1], duration=0.2)
                    pyautogui.click()
                    time.sleep(0.4)
                    
            def escribe(text):
                if text:
                    # Usamos PORTAPAPELES (Paste) en vez de typewrite para no fallar ni una minuscula ni dejarse letras ("si iene")
                    pyperclip.copy(str(text))
                    time.sleep(0.1)
                    pyautogui.hotkey('ctrl', 'v')
                    time.sleep(0.3)
                    
            def clica_y_escribe(key, text):
                clica(key)
                escribe(text)
                
            # 1. Fecha recepcion - CLICANDO EN CALENDARIO COMO PIDE EL USUARIO
            clica("Fecha recepción")
            time.sleep(1.5) # Aparece popup rojo de Albaran
            pyautogui.press('esc') # Lo cerramos
            time.sleep(0.8)
            
            # El usuario ha pedido estrictamente CLICAR en el día del calendario 
            # Utilizará la coordenada que ha guardado en la calibración
            if "Día de hoy (Calendario)" in c:
                clica("Día de hoy (Calendario)")
                time.sleep(0.5)
            
            # 2. Técnico Borrador
            clica_y_escribe("Técnico Borrador", "07")
            time.sleep(1.0)
            pyautogui.press('enter') # ELIMINADO EL DOWN (flecha abajo)
            time.sleep(0.5)
            
            # 3. Recepcionista
            clica_y_escribe("Recepcionista", "07")
            time.sleep(1.0)
            pyautogui.press('enter')
            time.sleep(0.5)
            
            # 4. Cliente
            clica_y_escribe("Cliente", "0000")
            time.sleep(1.5)
            pyautogui.press('enter') 
            time.sleep(0.5)
            
            # 5. Marca
            clica_y_escribe("Marca", "55")
            time.sleep(1.0)
            pyautogui.press('enter')
            time.sleep(0.5)
            
            # 6. Modelo
            clica_y_escribe("Modelo", modelo)
            time.sleep(1.5)
            pyautogui.press('enter')
            time.sleep(0.5)
            
            # 7. Número de serie
            clica_y_escribe("Número de serie", sn)
            
            # 8. Tipo de servicio - EL PRIMER HUECO ESTÁ EN BLANCO
            clica("Tipo de servicio")
            time.sleep(0.5)
            pyautogui.press('down') # PARA BAJAR A GARANTIA
            time.sleep(0.3)
            pyautogui.press('enter')
            time.sleep(0.5)
            
            # 9. Garantía
            clica("Garantía")
            time.sleep(0.3)
            
            # 10. Fecha de compra
            clica_y_escribe("Fecha de compra", fecha_compra)
            pyautogui.press('tab')
            time.sleep(0.3)
            
            # 11. Su referencia
            clica_y_escribe("Su referencia", su_ref)
            
            # 12. Accesorios POPUP
            if "Botón ?" in c:
                clica("Botón ?")
                time.sleep(2.5)
                
                if "Botón Usado (Popup)" in c:
                    clica("Botón Usado (Popup)")
                    time.sleep(0.5) 
                    
                if "Obs Generales (Popup)" in c:
                    clica_y_escribe("Obs Generales (Popup)", obs_accesorios)
                    
                if "Guardar (Popup)" in c:
                    clica("Guardar (Popup)")
                    time.sleep(1.0)
                    for _ in range(5):
                        pyautogui.press('esc')
                        time.sleep(0.2)
                    time.sleep(1.0)
                    
            # 13. Observaciones usuario (al final)
            clica_y_escribe("Observaciones usuario", averia)
            pyautogui.press('tab')
            
            messagebox.showinfo("Fin", "¡Terminado con éxito!")
        except Exception as e:
            messagebox.showerror("Error en el Robot", str(e))
        finally:
            self.root.deiconify()

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
