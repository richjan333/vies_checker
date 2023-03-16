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
from zeep import CachingClient, Transport


class ViesChecker:
    def __init__(self):

        self.queue = Queue()
        self.max_row = None
        self.sheet = None
        self.wb = None
        self.error_queue = Queue()
        self.running_queue = Queue()
        self.vat_list = []

        self.root = tk.Tk()
        self.root.title("VAT number checker")
        self.root.geometry("400x800")
        self.pick_file = tk.Button(self.root, text="Vyber soubor", command=self.__file_path_fce, height=2, width=10)
        self.pick_file.pack(padx=10, pady=10)
        self.file_path = tk.Label(self.root, text="Zde bude vypsána cesta k souboru.", bg="red", width=50)
        self.file_path.pack(padx=10, pady=10)

        self.start = tk.Button(self.root, text="Start", command=self.work_flow, height=2, width=10)
        self.start.pack(padx=10, pady=10)

        self.status_bar = tk.Label(self.root, text="Status: ", bg="white", width=50)
        self.status_bar.pack()
        self.status_bar_threads = tk.Label(self.root, text="Zpracováno: 0/0 vláken.", bg="white", width=50)
        self.status_bar_threads.pack()

        self.error_counter_label = tk.Label(self.root, text="Počet chyb: ", bg="white", width=50)
        self.error_counter_label.pack()

        list_frame = tk.Frame(self.root)
        list_frame.pack()
        left_frame = tk.Frame(list_frame)
        left_frame.pack(side=tk.LEFT)
        tk.Label(left_frame, text="Chyby").pack()
        self.error_list = tk.Listbox(left_frame, width=20, height=30)
        self.error_list.pack()

        right_frame = tk.Frame(list_frame)
        right_frame.pack(side=tk.RIGHT)
        tk.Label(right_frame, text="Ověřené DIČ.").pack()
        self.ok_list = tk.Listbox(right_frame, width=20, height=30)
        self.ok_list.pack()

        # if there is a config.txt file, it will be used as a default path
        try:
            with open("config.txt", "r") as f:
                self.file_path.config(text=f.read(), bg="green")
                self.__load_file()
        except FileNotFoundError:
            pass
        self.root.mainloop()

    def __load_file(self):
        # function for loading file
        self.wb = openpyxl.load_workbook(self.file_path.cget("text"))
        self.sheet = self.wb.active
        self.max_row = self.sheet.max_row
        self.vat_list = []
        for i in range(1, self.max_row + 1):
            # if second column is empty, it will be checked
            if self.sheet.cell(row=i, column=2).value is None:
                cell = self.sheet.cell(row=i, column=1)
                self.vat_list.append(cell.value)
        self.status_bar_threads.config(text="Zpracováno: " + str(0) + "/" + str(len(self.vat_list)))
        self.status_bar.config(text="Status: Soubor načten.")
        self.start.config(state=tk.NORMAL)

    def __file_path_fce(self):
        path = filedialog.askopenfilename()
        self.file_path.config(text=path, bg="green")
        with open("config.txt", "w") as f:
            f.write(path)
            self.__load_file()

    def work_flow(self):
        """
        Main function for checking VAT numbers
        :return:
        """
        # disable buttons
        self.pick_file.config(state=tk.DISABLED)
        self.start.config(state=tk.DISABLED)
        # create list of threads
        threads = []
        for i in self.vat_list:
            status = [False, 3]
            thread = threading.Thread(target=check_vat,
                                      args=(i, self.queue, self.error_queue, self.running_queue, status))
            threads.append([thread, i, status])
            print("Thread " + i + " initialized")

        finished = 0

        while finished < len(threads):
            start_time = time.time()
            threads[finished][0].start()
            wait_time = 30
            while threads[finished][2][0] is False and threads[finished][2][
                1] > 0 and time.time() - start_time < wait_time:
                left_count = threads[finished][2][1]
                if left_count < 3 or int(wait_time - (time.time() - start_time)) < 20:
                    self.status_bar.config(background="orange")
                else:
                    self.status_bar.config(background="white")
                self.status_bar.config(text="Čekám na " + threads[finished][1] + "\n Počet zbývajících pokusů: " +
                                            str(left_count) + "\n Čas do timeoutu: " + str(
                    int(wait_time - (time.time() - start_time))) + "s")
                time.sleep(0.3)
                self.root.update()

            # add to listbox
            if threads[finished][2][0] is False:
                self.error_list.insert(0, threads[finished][1])
                self.error_counter_label.config(text="Počet chyb: " + str(self.error_list.size()))
            else:
                self.ok_list.insert(0, threads[finished][1])
            time.sleep(1)
            if threads[finished][2][0] is False:
                threads[finished][0].join(timeout=10)
            finished += 1
            self.status_bar_threads.config(text="Zpracováno: " + str(finished) + "/" + str(len(threads)))
            self.root.update()

            self.save()
        self.status_bar.config(text="Status: Dokončeno.")
        self.pick_file.config(state=tk.NORMAL)


    def save(self):
        while not self.queue.empty():
            vat, result = self.queue.get()

            for i in range(1, self.max_row + 1):
                if self.sheet.cell(row=i, column=1).value == vat:
                    if result:
                        self.sheet.cell(row=i, column=2).value = "VALID"
                    else:
                        self.sheet.cell(row=i, column=2).value = "INVALID"
        while not self.error_queue.empty():
            vat = self.error_queue.get()
            for i in range(1, self.max_row + 1):
                if self.sheet.cell(row=i, column=1).value == vat:
                    self.sheet.cell(row=i, column=2).value = "ERROR"
        # save file to new file
        self.wb.save(self.file_path.cget("text"))


def check_vat(vat, queue, error_queue, running_queue, status_array=None):
    if status_array is None:
        status_array = [False, 3]
    done = status_array[0]

    try:
        while status_array[1] > 0:
            try:
                running_queue.put(vat)
                transport = Transport(timeout=5)
                client = CachingClient('http://ec.europa.eu/taxation_customs/vies/checkVatService.wsdl',
                                       transport=transport).service
                result = client.checkVat(countryCode=vat[:2], vatNumber=vat[2:])
                queue.put((vat, result["valid"]))
                error_counter = 0
                status_array[0] = True
                break
            except Exception as e:
                status_array[1] -= 1
                logging.error(e)
                time.sleep(3)
                running_queue.get()
        if not status_array[0]:
            error_queue.put(vat)

    except Exception as e:
        error_queue.put(vat)
        print(e)
    finally:
        running_queue.get()
        print("Thread " + vat + " finished")

        # print active threads


def wait_until_clear(queue, index=None, threads=None):
    start_time = time.time()
    suffix = "/"
    while running_queue.qsize() >= 3:
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
        if not index and not threads:
            break
        else:
            if start_time + 30 < time.time():
                print("Thread " + str(index) + " is stuck, killing it.")
                error_queue.put(threads[index - 1][1])
                threads[index - 1][0].join(timeout=5)
                break


def solver_fce():
    # create queue
    queue = Queue()
    # create list of VAT numbers

    active_threads = 0
    finished = 0

    for index, t in enumerate(threads):
        wait_until_clear(queue=queue, index=index, threads=threads)
        if running_queue.qsize() < 3:
            try:
                t[0].start()
                active_threads += 1
            except Exception as e:
                print(e)
                logging.error(e)

    wait_until_clear(queue=queue)

    status_bar.config(text="Dokončeno. Ukládám do souboru.")
    status_bar_threads.config(
        text="Správně zpracovaných dotazů: " + str(queue.qsize()) + "/" + str(len(vat_list)) + " DIČ.")
    error_counter_label.config(text="Počet chyb: " + str(error_queue.qsize()))
    root.update()
    save(queue=queue, max_row=max_row, sheet=sheet, wb=wb)


def main_loop(error_queue, queue):
    pass


if __name__ == '__main__':
    ViesChecker()
