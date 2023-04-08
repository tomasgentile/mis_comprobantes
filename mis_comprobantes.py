"""
Descarga archivos (y los concatena) del servicio de Mis comprobantes de la web de AFIP
"""
import os
from pathlib import Path
import shutil
import time
import glob
from tkinter import Tk, ttk, Frame, Label, Entry, StringVar, Radiobutton, W, E
from datetime import datetime
import threading
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException, WebDriverException
import pandas as pd


class App:
    """ Inicia la pantalla y recibe los valores para la busqueda """
    def __init__(self):
        self.window = Tk()
        self.window.eval('tk::PlaceWindow . center')
        self.window.title('Mis comprobantes')

        # Creating a frame container
        frame = Frame(self.window)
        frame.grid(row=0, column=0, columnspan=3, pady=20, ipadx=10, ipady=10)

        # CUIT input
        Label(frame, text='CUIT Login: ', justify='left', anchor='w').grid(row=1, column=0, padx=(10, 5), sticky=W)
        self.cuit_login = Entry(frame)
        self.cuit_login.focus()
        self.cuit_login.grid(row=1, column=1)

        # CUIT contribuyente input
        Label(frame, text='CUIT Contribuyente: ', justify='left', anchor='w').grid(row=2, column=0, padx=(10, 5), sticky=W)
        self.cuit_cont = Entry(frame)
        self.cuit_cont.grid(row=2, column=1)

        # Password input
        Label(frame, text='Clave: ', justify='left', anchor='w').grid(row=3, column=0, padx=(10, 5), sticky=W)
        self.password = Entry(frame)
        self.password.grid(row=3, column=1)

        # Period input
        Label(frame, text='mm').grid(row=4, column=1)
        Label(frame, text='yyyy').grid(row=4, column=2)
        Label(frame, text='Periodo desde: ', justify='left', anchor='w').grid(row=5, column=0, padx=(10, 5), sticky=W)
        Label(frame, text='Periodo hasta: ', justify='left', anchor='w').grid(row=6, column=0, padx=(10, 5),sticky=W)
        self.month_from = Entry(frame, justify='center')
        self.month_from.grid(row=5, column=1)
        self.year_from = Entry(frame, justify='center')
        self.year_from.grid(row=5, column=2)

        self.month_to = Entry(frame, justify='center')
        self.month_to.grid(row=6, column=1)
        self.year_to = Entry(frame, justify='center')
        self.year_to.grid(row=6, column=2)
        self.period = []

        # Emitidos / Recibidos
        self.comp_type = StringVar()
        self.comp_type.set('emitidos')
        Radiobutton(frame, text='Emitidos', variable=self.comp_type, value='emitidos').grid(row=7, column=0)
        Radiobutton(frame, text='Recibidos', variable=self.comp_type, value='recibidos').grid(row=7, column=1)

        # Output messages
        self.messsage_text = ''
        self.message = Label(frame)
        # Submit button
        ttk.Button(frame, text='Buscar', command=self.start).grid(row=8, columnspan=3, padx=(15, 0), pady=10, sticky=W + E)

        # Download folder
        self.parent_folder = os.path.expanduser('~') + '\\Downloads' + '\\MisComprobantes'
        self.download_folder = ''
        self.update_messages()
        self.window.mainloop()


    def update_messages(self):
        """ Actualiza el label con mesajes """
        self.message.config(text=self.messsage_text)
        self.message.grid(row=9, column=0, columnspan=3, padx=10)
        self.message.after(1, self.update_messages)


    def start(self):
        """
        Abre un thread para la busqueda
        """
        threading.Thread(target=self.search).start()

    def validation(self):
        """ Valida que los datos esten completos, que los meses esten entre 1 y 12, que los años sean mayores a 2000 y que
        no sean mayores al año actual, y que la fecha desde sea anterior a la fecha hasta """
        try:
            # Date completed
            if len(self.cuit_login.get().strip()) == 11 and len(self.password.get().strip()) != 0 and len(self.cuit_cont.get().strip()) == 11 and len(self.month_from.get()) != 0 and len(self.year_from.get()) != 0 and len(self.month_to.get()) != 0 and len(self.year_to.get()) != 0:
                # Validate Months
                if int(self.month_from.get()) > 0 and int(self.month_from.get()) <= 12 and int(self.month_to.get()) > 0 and int(self.month_to.get()):
                    current_year = datetime.now().year
                    # Validate year
                    if int(self.year_from.get()) > 2000 and int(self.year_from.get()) <= current_year and int(self.year_to.get()) > 0 and int(self.year_to.get()) <= current_year:
                        if int(self.year_from.get()) <= int(self.year_to.get()):
                            if int(self.year_from.get()) == int(self.year_to.get()):
                                if int(self.month_from.get()) <= int(self.month_to.get()):
                                    return True
                            else:
                                return True
        except ValueError:
            return False
    
    def generate_month_str(self, mth):
        """ Recibe un entero y devuelve un string, adicionando 0 si corresponde """
        if mth < 10:
            output = f'0{mth}'
        else:
            output = str(mth)
        return output


    def search(self):
        """ Realiza la busqueda de acuerdo a los parametros recibidos y descarga los archivos """
        if self.validation():
            self.period = []
            self.messsage_text = 'Iniciando sesión en AFIP'
            self.message.config(fg='green')

            # Search period
            if int(self.year_from.get()) == int(self.year_to.get()):
            # Si el periodo inicia y finaliza en el mismo año
                for month_idx in range(int(self.month_from.get()), int(self.month_to.get()) + 1):
                    self.period.append([self.generate_month_str(month_idx), str(self.year_to.get())])
            elif int(self.year_to.get()) - int(self.year_from.get()) == 1:
                # Si el periodo va de un año a otro
                for month_idx in range(int(self.month_from.get()), 13):
                    self.period.append([self.generate_month_str(month_idx), str(self.year_from.get())])
                for month_idx in range(1, int(self.month_to.get()) + 1):
                    self.period.append([self.generate_month_str(month_idx), str(self.year_to.get())])
            else:
                # Si el periodo abarca varios años
                for month_idx in range(int(self.month_from.get()), 13):
                    self.period.append([self.generate_month_str(month_idx), str(self.year_from.get())])
                for year_idx in range(int(self.year_from.get()) + 1, int(self.year_to.get())):
                    for month_idx in range(1, 13):
                        self.period.append([self.generate_month_str(month_idx), str(year_idx)])
                for month_idx in range(1, int(self.month_to.get())):
                    self.period.append([self.generate_month_str(month_idx), str(self.year_to.get())])

            # Download folder
            Path(self.parent_folder).mkdir(exist_ok=True)
            os.chdir(self.parent_folder)
            Path(self.cuit_cont.get().strip()).mkdir(exist_ok=True)
            self.download_folder = os.path.join(self.parent_folder, self.cuit_cont.get().strip(), 'Temp')
            # Opciones de Webdriver
            options = webdriver.ChromeOptions()
            prefs = {'download.default_directory': self.download_folder}
            options.add_argument("--window-size=1920,1080")
            options.add_argument("--start-maximized")
            options.add_experimental_option("prefs", prefs)
            options.add_experimental_option("excludeSwitches", ["enable-logging"])

            try:
                chrome_browser = webdriver.Chrome(options=options)
                chrome_browser.implicitly_wait(20)
                chrome_browser.get('https://auth.afip.gob.ar/contribuyente_/login.xhtml')
                chrome_browser.maximize_window()
            except WebDriverException:
                self.messsage_text = 'Error al cargar la página de AFIP, intente más tarde'
                return

            # Setup wait for later
            wait = WebDriverWait(chrome_browser, 10)

            try:
            # Login
                cuit_input = chrome_browser.find_element(By.ID, 'F1:username')
                cuit_input.send_keys(self.cuit_login.get().strip())
                next_button = chrome_browser.find_element(By.ID, 'F1:btnSiguiente')
                next_button.click()
                password_input = chrome_browser.find_element(By.ID, 'F1:password')
                password_input.send_keys(self.password.get().strip())
                login_button = chrome_browser.find_element(By.ID, 'F1:btnIngresar')
                login_button.click()
            except NoSuchElementException:
                self.messsage_text = 'Error al intentar loguearse a la página de AFIP'
                self.message.config(fg='red')
                chrome_browser.quit()
                return
            
            try:
                self.messsage_text = 'Iniciando descarga'
                self.message.config(fg='green')
                # Panel General
                time.sleep(1)
                service_input = chrome_browser.find_element(By.ID, 'buscadorInput')
                service_input.send_keys('Mis comprobantes')
                search_box = chrome_browser.find_element(By.ID, 'resBusqueda')
                search_box.click()

                # Mis comprobantes
                wait.until(expected_conditions.number_of_windows_to_be(2))
                mis_comprobantes_window = chrome_browser.window_handles[1]
                chrome_browser.switch_to.window(mis_comprobantes_window)

                # Elegir contribuyente
                select_contrib = chrome_browser.find_element(By.TAG_NAME, 'h1')
                if select_contrib.text == 'Elegí una persona para ingresar':
                    contrib_button = chrome_browser.find_element(
                        By.XPATH, '/html/body/form/main/div/div/div[2]/div/a/div/div[2]/p')
                    text_button = contrib_button.text.replace('-', '')
                    if text_button == self.cuit_cont.get().strip():
                        contrib_button.click()
                    else:
                        represented_users = chrome_browser.find_elements(By.TAG_NAME, 'small')
                        for user in represented_users:
                            if user.text.replace('-', '') == self.cuit_cont.get().strip():
                                user.click()
                                break

                emitidos_btn = chrome_browser.find_element(By.ID, 'btnEmitidos')
                recibidos_btn = chrome_browser.find_element(By.ID, 'btnRecibidos')

                if self.comp_type.get() == 'emitidos':
                    emitidos_btn.click()
                else:
                    recibidos_btn.click()

                for item in self.period:
                    month = item[0]
                    year = item[1]
                    date_from = f'01/{month}/{year}'
                    date_to = ''

                    if month == '02':
                        date_to = f'28/{month}/{year}'
                    elif month in ('04', '06', '09', '11'):
                        date_to = f'30/{month}/{year}'
                    else:
                        date_to = f'31/{month}/{year}'

                    self.messsage_text = f'Descargando {month}/{year}'
                    self.message.config(fg='green')

                    fecha_emision_input = chrome_browser.find_element(By.ID, 'fechaEmision')
                    fecha_emision_input.clear()
                    fecha_emision_input.send_keys(f'{date_from} - {date_to}')
                    apply_btn = chrome_browser.find_element(By.CSS_SELECTOR, '.applyBtn')
                    apply_btn.click()
                    submit_btn = chrome_browser.find_element(By.ID, 'buscarComprobantes')
                    submit_btn.click()

                    # Descargar archivo
                    wait.until(expected_conditions.element_to_be_clickable((By.CLASS_NAME, 'buttons-csv'))).click()
                    time.sleep(1)

                    consult_btn = chrome_browser.find_element(
                        By.XPATH, '//*[@id="tabsComprobantes"]/li[1]/a')
                    consult_btn.click()

                self.concatenate_files()
                chrome_browser.quit()
            except ElementNotInteractableException:
                self.messsage_text = 'Hubo un error en la descarga, intente más tarde'
                chrome_browser.quit()
                return

        else:
            self.message.config(fg='red')
            self.messsage_text  = 'Verifique los datos ingresados'

    def concatenate_files(self):
        """ Controla que haya archivos descargados y los concatena """

        files = glob.glob(f'{self.download_folder}\\*.csv')
        if len(files) == 0:
            self.message.config(fg='red')
            self.messsage_text = 'Error al realizar la descarga, intente más tarde'
            return
        # Ordena los archivos por fecha de creación
        files.sort(key=os.path.getctime)

        file_list = []

        for file in files:
            data = pd.read_csv(file)
            file_list.append(data)

        data_frame = pd.concat(file_list, ignore_index=True)
        cuit_folder = os.path.join(self.parent_folder, self.cuit_cont.get().strip())
        os.chdir(cuit_folder)
        excel_writer = pd.ExcelWriter(f'Comprobantes {self.comp_type.get()} CUIT {self.cuit_cont.get().strip()} - {self.month_from.get()}-{self.year_from.get()} - {self.month_to.get()}-{self.year_to.get()}.xlsx')  # pylint: disable=abstract-class-instantiated
        data_frame.to_excel(excel_writer)
        excel_writer.save()
        shutil.rmtree(self.download_folder)
        self.messsage_text = f'Descarga disponible en: {cuit_folder}'


if __name__ == '__main__':
    GIU = App()
