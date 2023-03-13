# project for checking VIES VAT number from excell file and results writes to second column of the same file
# uses tkinter, openpyxl, zeep
# results are returned to queue and then written to file
import logging
import threading
import time
import tkinter as tk
from queue import Queue
from tkinter import filedialog
import openpyxl
from zeep import Client

error_queue = Queue()
running_queue = Queue()
vat_list = Queue()


# function for checking VAT number
def check_vat(vat, queue):
    try:
        error_counter = 3
        done = False
        while error_counter > 0:
            try:
                running_queue.put(vat)
                client = Client('http://ec.europa.eu/taxation_customs/vies/checkVatService.wsdl')
                result = client.service.checkVat(countryCode=vat[:2], vatNumber=vat[2:])
                queue.put((vat, result["valid"]))
                error_counter = 0
                done = True
                running_queue.get()
            except Exception as e:
                error_counter -= 1
                logging.error(e)
                time.sleep(3)
        if not done:
            error_queue.put(vat)
    except Exception as e:
        error_queue.put(vat)
        print(e)
        logging.error(e)


def solver_fce():
    # import data from file
    wb = openpyxl.load_workbook(file_path.cget("text"))
    sheet = wb.active
    max_row = sheet.max_row
    # create queue
    queue = Queue()
    # create list of VAT numbers
    vat_list = []
    for i in range(1, max_row + 1):
        cell = sheet.cell(row=i, column=1)
        vat_list.append(cell.value)

    # create list of threads
    threads = []
    for i in vat_list:
        thread = threading.Thread(target=check_vat, args=(i, queue))
        threads.append((thread, i))
    active_threads = 0
    finished = 0
    suffix = "/"
    for t in threads:
        while running_queue.qsize() >= 5:
            list_of_running_vats = list(running_queue.queue)
            list_of_running_vats_text = ""
            for i in list_of_running_vats:
                list_of_running_vats_text += i + ", \n"
            status_bar.config(text="Aktuálně zpracovávám: \n" + str(list_of_running_vats_text) + " DIČ." + suffix)
            status_bar_threads.config(
                text="Správně zpracovaných dotazů: " + str(queue.qsize()) + "/" + str(len(vat_list)) + " DIČ.")
            error_counter_label.config(text="Počet chyb: " + str(error_queue.qsize()))
            time.sleep(0.3)
            if suffix == "/":
                suffix = "-"
            elif suffix == "-":
                suffix = "\\"
            elif suffix == "\\":
                suffix = "|"
            elif suffix == "|":
                suffix = "/"

            if not error_queue.empty():
                error_list.insert(tk.END, error_queue.get())
            if not queue.empty():
                ok_list_list = list(queue.queue)
                ok_list.delete(0, tk.END)
                for i in ok_list_list:
                    ok_list.insert(0, i[0])
            root.update()
        if running_queue.qsize() <= 5:
            t[0].start()
            active_threads += 1

    while running_queue.qsize() > 0:
        list_of_running_vats = list(running_queue.queue)
        list_of_running_vats_text = ""
        for i in list_of_running_vats:
            list_of_running_vats_text += i + ", \n"
        status_bar.config(text="Aktuálně zpracovávám: \n" + str(list_of_running_vats_text) + " DIČ." + suffix)
        status_bar_threads.config(
            text="Správně zpracovaných dotazů: " + str(queue.qsize()) + "/" + str(len(vat_list)) + " DIČ.")
        error_counter_label.config(text="Počet chyb: " + str(error_queue.qsize()))
        time.sleep(0.3)
        if suffix == "/":
            suffix = "-"
        elif suffix == "-":
            suffix = "\\"
        elif suffix == "\\":
            suffix = "|"
        elif suffix == "|":
            suffix = "/"

        if not error_queue.empty():
            error_list.insert(tk.END, error_queue.get())
        if not queue.empty():
            ok_list_list = list(queue.queue)
            ok_list.delete(0, tk.END)
            for i in ok_list_list:
                ok_list.insert(0, i[0])
        root.update()
    # write results to file

    status_bar.config(text="Dokončeno. Ukládám do souboru.")
    status_bar_threads.config(
        text="Správně zpracovaných dotazů: " + str(queue.qsize()) + "/" + str(len(vat_list)) + " DIČ.")
    error_counter_label.config(text="Počet chyb: " + str(error_queue.qsize()))
    root.update()
    while not queue.empty():
        vat, result = queue.get()
        for i in range(2, max_row + 1):
            if sheet.cell(row=i, column=1).value == vat:
                if result:
                    sheet.cell(row=i, column=2).value = "VALID"
                else:
                    sheet.cell(row=i, column=2).value = "INVALID"
    while not error_queue.empty():
        vat = error_queue.get()
        for i in range(2, max_row + 1):
            if sheet.cell(row=i, column=1).value == vat:
                sheet.cell(row=i, column=2).value = "ERROR"
    # save file to new file
    wb.save("output.xlsx")


def main_loop(error_queue, queue):
    pass


def file_path_fce():
    path = filedialog.askopenfilename()
    file_path.config(text=path, bg="green")
    with open("config.txt", "w") as f:
        f.write(path)


root = tk.Tk()
root.title("VAT number checker")
root.geometry("400x800")
pick_file = tk.Button(root, text="Vyber soubor", command=file_path_fce, height=2, width=10)
pick_file.pack(padx=10, pady=10)
file_path = tk.Label(root, text="Zde bude vypsána cesta k souboru.", bg="red", width=50)
file_path.pack(padx=10, pady=10)

start = tk.Button(root, text="Start", command=solver_fce, height=2, width=10)
start.pack(padx=10, pady=10)

status_bar = tk.Label(root, text="Status: ", bg="white", width=50)
status_bar.pack()
status_bar_threads = tk.Label(root, text="Zpracováno: 0/0 vláken.", bg="white", width=50)
status_bar_threads.pack()

error_counter_label = tk.Label(root, text="Počet chyb: ", bg="white", width=50)
error_counter_label.pack()

list_frame = tk.Frame(root)
list_frame.pack()
left_frame = tk.Frame(list_frame)
left_frame.pack(side=tk.LEFT)
tk.Label(left_frame, text="Chyby").pack()
error_list = tk.Listbox(left_frame, width=20, height=30)
error_list.pack()

right_frame = tk.Frame(list_frame)
right_frame.pack(side=tk.RIGHT)
tk.Label(right_frame, text="Ověřené DIČ.").pack()
ok_list = tk.Listbox(right_frame, width=20, height=30)
ok_list.pack()

# if there is a config.txt file, it will be used as a default path
try:
    with open("config.txt", "r") as f:
        file_path.config(text=f.read(), bg="green")
except FileNotFoundError:
    pass
root.mainloop()
