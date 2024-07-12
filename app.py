import customtkinter as ctk
from tkinter import messagebox

from paquete.envios_pre import envio_correos_previo
from paquete.envios_post import envio_correos_post
from paquete.creacion_excel import crear_nuevo_excel
from users.emails_permitidos import lista_emails

# ----- Crear la ventana principal --------
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Producción")
        self.root.geometry("400x300") 

        ctk.CTkLabel(root, text="Email:").pack(pady=5)
        self.entry_email = ctk.CTkEntry(root)
        self.entry_email.pack(pady=5)

        ctk.CTkLabel(root, text="Contraseña:").pack(pady=5)
        self.entry_contraseña = ctk.CTkEntry(root, show="*")
        self.entry_contraseña.pack(pady=5)

        ctk.CTkLabel(root, text="Tipo de mensaje:").pack(pady=5)
        self.opciones_msj = ctk.StringVar(value="Seleccionar mensaje")
        self.combobox_msj = ctk.CTkOptionMenu(root, variable=self.opciones_msj, values=["Previo", "Posterior"])
        self.combobox_msj.pack(pady=5)

        print(self.combobox_msj)
        ctk.CTkButton(root, text="ENVIAR", command=self.ejecucion).pack(pady=20)

# ------------ Validación --------------------

    def validar_email(self):
        email_get = self.entry_email.get()

        for user_info in lista_emails:
            if email_get == user_info["email"]:
                print("Éxito", "El email coincide")
                self.usuario = user_info["user"]
                self.email_validado = user_info["email"]
                return True
        
        messagebox.showinfo("Error", "El email no se encuentra en la lista")
        return False
    
# -------------- Ejecución ----------------------

    def ejecucion(self):
        if self.validar_email():

            tipo_mensaje = self.opciones_msj.get()
            print(tipo_mensaje)

            if tipo_mensaje == "Previo":
                confirmacion = messagebox.askyesno(message="Está por enviar los correos PREVIOS a la fecha de vencimiento ¿Desea continuar?", title="RECORDATORIO")
                print(confirmacion)
                if confirmacion:
                    envio_correos_previo(self.email_validado, self.entry_contraseña.get(), self.usuario)
                    
            elif tipo_mensaje == "Posterior":
                confirmacion = messagebox.askyesno(message="Previamente debio actualizar el excel 'produccion' con las compañias que no presentaron ¿Desea continuar?", title="RECORDATORIO")
                print(confirmacion)
                if confirmacion:
                    crear_nuevo_excel()
                    envio_correos_post(self.email_validado, self.entry_contraseña.get(), self.usuario)

            else:
                messagebox.showinfo("Error", "Debe seleccionar un tipo de mensaje")

if __name__ == "__main__":
    ctk.set_appearance_mode("dark") 
    ctk.set_default_color_theme("blue")  

    root = ctk.CTk()
    app = App(root)
    root.mainloop()
