from datetime import datetime
from PyQt5 import uic
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import QIntValidator
from db_config import *
import xlsxwriter
import sys
import os


class ConfigureScreen:
    # menampilkan widget ke tengan
    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    # Mengambil posisi widget ketika mouse ditekan
    def mousePressEvent(self, event):
        self.oldPos = event.globalPos()

    # Memindahkan posisi widget ketika mouse digeser
    def mouseMoveEvent(self, event):
        delta = QPoint(event.globalPos() - self.oldPos)
        self.move(self.x() + delta.x(), self.y() + delta.y())
        self.oldPos = event.globalPos()


class LoginWindow(QWidget, ConfigureScreen):
    def __init__(self):
        super(LoginWindow, self).__init__()
        uic.loadUi('login.ui', self)

        # membuat widget menjadi frameless
        self.setWindowFlags(Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)

        # Konfigurasi interaksi push button dengan method yang ada
        self.pushButtonLogin.clicked.connect(lambda: self.login())
        self.pushButtonClose.clicked.connect(lambda: self.close_menu())

    # Membuat method untuk login
    def login(self):
        # Melakukan pengecekan apakah input username dan password kosong atau tidak
        if self.lineEditUsername.text() != "" or self.lineEditPassword.text() != "":
            # melakukan pengecekan apakah username dan password yang diisi adalah "admin"
            if self.lineEditUsername.text() == "admin" and self.lineEditPassword.text() == "admin":
                # Membuka window mainMenu
                self.mainMenu = MainMenu()
                self.mainMenu.show()
                self.close()
                QMessageBox.about(self, 'Success', 'Login Berhasil! Selamat Datang')
            else:
                QMessageBox.about(self, 'Peringatan', 'Username atau Password salah')
        else:
            QMessageBox.about(self, 'Peringatan', 'Username dan Password tidak boleh kosong')

    # Membuat method untuk menutup widget
    def close_menu(self):
        self.close()


class MainMenu(QWidget, ConfigureScreen):
    def __init__(self):
        super(MainMenu, self).__init__()
        uic.loadUi('mainMenu.ui', self)

        # membuat widget menjadi frameless
        self.setWindowFlags(Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)

        # Inisialisasi attribute karyawan dengan value class Karyawan
        self.karyawan = Karyawan()

        # Set validasi agar lineEditLembur dan lineEditNip hanya bisa inputkan dengan angka
        self.lineEditLembur.setValidator(QIntValidator(self))
        self.lineEditNip.setValidator(QIntValidator(self))

        # Konfigurasi interaksi push button dengan method yang ada
        self.pushButtonLogout.clicked.connect(lambda: self.logout())
        self.pushButtonClose.clicked.connect(lambda: self.close_menu())
        self.pushButtonReset.clicked.connect(lambda: self.reset())
        self.pushButtonTambah.clicked.connect(lambda: self.store_karyawan())
        self.pushButtonUbah.clicked.connect(lambda: self.update_karyawan())
        self.pushButtonHapus.clicked.connect(lambda: self.delete_karyawan())
        self.pushButtonExport.clicked.connect(lambda: self.export_data_karyawan())
        self.pushButtonCetak.clicked.connect(lambda: self.export_slip_gaji())

        # Konfigurasi interaksi table, comboBox, dan lineedit dengan method yang ada
        self.tableKaryawan.itemClicked.connect(lambda: self.set_input_when_clicked())
        self.comboBoxJabatan.currentIndexChanged.connect(lambda: self.show_total_gaji())
        self.lineEditLembur.textChanged.connect(lambda: self.show_total_gaji())

        # menjalankan method untuk menampilkan karyawan dari database ke tabel
        self.show_karyawan_to_table()

    # Method untuk menghitung total gaji berdasarkan jabatan, gaji , tunjangan, ppn, dan uang lembur
    def calculate(self):
        # Menginisialisasikan variable gaji_dict dengan dictionary yang berisi jabatan, gaji, tunjangan, dan ppn
        gaji_dict = {
            'direktur': {
                'gaji': 20000000,
                'tunjangan': 5000000,
                'ppn': 2000000,
            },
            'manager': {
                'gaji': 10000000,
                'tunjangan': 1800000,
                'ppn': 1000000,
            },
            'karyawan': {
                'gaji': 5000000,
                'tunjangan': 750000,
                'ppn': 500000
            }
        }

        # Memformat teks input dan mengkalkulasi total gaji
        jabatan = self.comboBoxJabatan.currentText().lower()
        penghasilan = gaji_dict[jabatan]
        uang_lembur = self.lineEditLembur.text() if self.lineEditLembur.text() != "" else 0
        calculate = (penghasilan['gaji'] + penghasilan['tunjangan'] - penghasilan['ppn']) + int(uang_lembur)

        # Mengembalikan nilai berupa list yang berisi list penghasilan dan total gaji
        return [penghasilan, calculate]

    # Method untuk mengisi inputan sesuai data yang di klik di tabel
    def set_input_when_clicked(self):
        row = self.tableKaryawan.currentRow()
        nip = self.tableKaryawan.item(row, 0)

        # Memanggil method dari class karyawan untuk mencari karyawan sesuai nip
        karyawan = self.karyawan.find_karyawan(nip.text())

        # Memasukkan data karyawan yang ditemukan ke dalam input dan label
        self.lineEditNip.setText(str(karyawan[1]))
        self.lineEditNama.setText(karyawan[2])
        self.comboBoxJabatan.setCurrentText(karyawan[3].capitalize())
        self.lineEditLembur.setText(str(karyawan[7]))
        self.show_total_gaji()

    # Method untuk menampilkan total gaji ke dalam label
    def show_total_gaji(self):
        # Menjalankan method calculate() untuk mendapatkan list penghasilan
        penghasilan = self.calculate()

        # menampilkan gaji, tunjangan dan ppn ke dalam input dan menampilkan total gaji ke dalam input
        self.lineEditGaji.setText(str(penghasilan[0]['gaji']))
        self.lineEditTunjangan.setText(str(penghasilan[0]['tunjangan']))
        self.lineEditPpn.setText(str(penghasilan[0]['ppn']))
        self.label_t_gaji.setText('Rp. {:,.2f}'.format(penghasilan[1]))

    # Method untuk memformat semua inputan menjadi tuple
    def input_to_tuple(self):
        # Menginisialisasikan variable data yang berisi semua input dalam bentuk tuple
        data = (
            self.lineEditNip.text(),
            self.lineEditNama.text(),
            self.comboBoxJabatan.currentText().lower(),
            self.lineEditLembur.text(),
            self.lineEditGaji.text(),
            self.lineEditTunjangan.text(),
            self.lineEditPpn.text()
        )

        # Mengembalikan data berupa inputan berbentuk tuple yang berada di variable data
        return data

    # Method untuk menampilkan karyawan ke dalam tabel
    def show_karyawan_to_table(self):
        # Memanggil method dari class karyawan untuk mendapatkan data karyawan
        karyawan = self.karyawan.get_karyawan()

        # memasukkan data karyawan ke dalam tabel
        self.tableKaryawan.setRowCount(0)
        for row, row_data in enumerate(karyawan):
            self.tableKaryawan.insertRow(row)
            for col, col_data in enumerate(row_data):
                if col > 2:
                    self.tableKaryawan.setItem(row, col, QTableWidgetItem('Rp. {:,.2f}'.format(col_data)))
                else:
                    self.tableKaryawan.setItem(row, col, QTableWidgetItem(str(col_data)))

    # Method untuk memasukkan karyawan ke dalam database
    def store_karyawan(self):
        nip = self.lineEditNip.text()
        nama = self.lineEditNama.text()

        # Mengecek apakah nip dan nama kosong atau tidak
        if nip == "" or nama == "":
            QMessageBox.about(self, 'Peringatan', 'NIP dan Nama tidak boleh kosong')
            return

        # Memanggil method dari class karyawan untuk mencari karyawan sesuai nip
        karyawan = self.karyawan.find_karyawan(nip)

        # Mengecek apakah data karyawan ada atau tidak
        if karyawan is not None:
            QMessageBox.about(self, 'Peringatan', 'NIP sudah digunakan')
            return

        # Menampilkan total gaji ke label ketika dijalankan
        self.show_total_gaji()
        data = self.input_to_tuple()

        # Memasukkan variable data yang berisi tuple ke argument store
        self.karyawan.store(data)
        self.show_karyawan_to_table()

        QMessageBox.about(self, 'Berhasil', 'Berhasil menambah data karyawan')

    # Method untuk mengubah data karyawan
    def update_karyawan(self):
        nip = self.lineEditNip.text()
        nama = self.lineEditNama.text()
        if nip == "" or nama == "":
            QMessageBox.about(self, 'Peringatan', 'NIP dan Nama tidak boleh kosong')
            return

        # Mengecek apakah nip karyawan ada atau tidak
        karyawan = self.karyawan.find_karyawan(nip)
        if karyawan is not None:
            data = self.input_to_tuple()

            # Memanggil method dari class karyawan untuk mengubah data karyawan sesuai nip
            self.karyawan.update(karyawan[0], data)
            self.show_karyawan_to_table()

            QMessageBox.about(self, 'Berhasil', 'Berhasil mengubah data karyawan')
        else:
            QMessageBox.about(self, 'Peringatan', 'Karyawan tidak ditemukan')

    # Method untuk menghapus karyawan
    def delete_karyawan(self):
        nip = self.lineEditNip.text()
        if nip == "":
            QMessageBox.about(self, 'Peringatan', 'Harap pilih user terlebih dahulu')
            return

        # Mengecek apakah nip karyawan ada atau tidak
        karyawan = self.karyawan.find_karyawan(nip)
        if karyawan is not None:
            self.karyawan.delete(karyawan[0])
            self.show_karyawan_to_table()

            QMessageBox.about(self, 'Berhasil', 'Berhasil menghapus data karyawan')
        else:
            QMessageBox.about(self, 'Peringatan', 'Karyawan tidak ditemukan')

    # Method untuk mengekspor semua data karyawan yang ada di database
    def export_data_karyawan(self):
        date = datetime.now()

        # Mengecek apakah folder bernama exported data ada atau tidak
        if not os.path.exists('exported_data'):
            os.makedirs('exported_data')

        # Membuat file excel dengan nama yang sudah ditentukan
        workbook = xlsxwriter.Workbook('exported_data/' + date.strftime("%Y%m%d%H%M%S") + '_exported_data.xlsx')
        worksheet = workbook.add_worksheet()

        # Konfigurasi sheet
        bold = workbook.add_format({'bold': True, 'border': 1})
        full_border = workbook.add_format({'border': 1})
        idr_format = workbook.add_format({'num_format': 'Rp #,##0.00', 'border': 1})
        worksheet.set_column(0, 7, 20)

        # Memasukkan data ke dalam excel
        worksheet.write(0, 0, "NIP", bold)
        worksheet.write(0, 1, "Nama", bold)
        worksheet.write(0, 2, "Jabatan", bold)
        worksheet.write(0, 3, "Gaji", bold)
        worksheet.write(0, 4, "Tunjangan", bold)
        worksheet.write(0, 5, "PPN", bold)
        worksheet.write(0, 6, "Uang Lembur", bold)
        worksheet.write(0, 7, "Total Gaji", bold)

        # Memasukkan data ke dalam berdasarkan data yang ada di database
        karyawan = self.karyawan.get_karyawan()
        for row, row_data in enumerate(karyawan):
            for col, col_data in enumerate(row_data):
                if col > 2:
                    if col == 6:
                        total_gaji = ((row_data[3] + row_data[4]) - row_data[5]) + row_data[6]
                        worksheet.write(row + 1, col + 1, total_gaji, idr_format)

                    worksheet.write(row + 1, col, col_data, idr_format)
                else:
                    worksheet.write(row + 1, col, str(col_data), full_border)

        workbook.close()
        QMessageBox.about(self, 'Berhasil', 'Data berhasil di export')

    # Method untuk mengeksport slip gaji dalam bentuk excel
    def export_slip_gaji(self):
        nip = self.lineEditNip.text()
        if nip == "":
            QMessageBox.about(self, 'Peringatan', 'Harap pilih user terlebih dahulu')
            return

        # Mengecek apakah nip karyawan ada atau tidak
        karyawan = self.karyawan.find_karyawan(nip)
        if karyawan is not None:
            # Mengecek apakah folder bernama exported data ada atau tidak
            if not os.path.exists('salary_slip'):
                os.makedirs('salary_slip')

            date = datetime.now()

            # Membuat file excel dengan nama yang sudah ditentukan
            workbook = xlsxwriter.Workbook(
                'salary_slip/' + date.strftime("%Y%m%d%H%M%S") + '_' +
                karyawan[2].replace(" ", "_") + str(karyawan[1]) + '.xlsx'
            )
            worksheet = workbook.add_worksheet()

            # Konfigurasi sheet
            idr_format = workbook.add_format(
                {'num_format': 'Rp #,##0.00', 'align': 'left', 'right': 1, 'bottom': 1, 'valign': 'vcenter', })
            data_border = workbook.add_format({'right': 1, 'bottom': 1, 'valign': 'vcenter', })
            label_border = workbook.add_format({'left': 1, 'bottom': 1, 'valign': 'vcenter', })
            header = workbook.add_format({
                'bold': True,
                'border': 1,
                'align': 'center',
                'valign': 'vcenter',
                'fg_color': '#c9c9c9',
            })

            # Hitung total gaji
            total_gaji = ((karyawan[4] + karyawan[5]) - karyawan[6]) + karyawan[7]

            # Konfigurasi ukuran baris dan kolom excel
            worksheet.set_default_row(20)
            worksheet.set_column(0, 0, 13)
            worksheet.set_column(1, 2, 10)
            worksheet.set_row(0, 25)
            worksheet.set_row(4, 25)

            # Memasukkan data ke dalam excel
            worksheet.write('A2', "NIP", label_border)
            worksheet.write('A3', "Nama", label_border)
            worksheet.write('A4', "Jabatan", label_border)
            worksheet.write('A6', "Gaji", label_border)
            worksheet.write('A7', "Tunjangan", label_border)
            worksheet.write('A8', "Uang Lembur", label_border)
            worksheet.write('A9', "PPN", label_border)
            worksheet.write('A10', "Total Gaji", label_border)

            # Memasukkan data ke dalam excel dengan merge sheet
            worksheet.merge_range("A1:C1", 'Salary Slip', header)
            worksheet.merge_range("A5:C5", 'Penerimaan', header)
            worksheet.merge_range("B2:C2", str(karyawan[1]), data_border)
            worksheet.merge_range("B3:C3", karyawan[2], data_border)
            worksheet.merge_range("B4:C4", karyawan[3].capitalize(), data_border)
            worksheet.merge_range("B6:C6", karyawan[4], idr_format)
            worksheet.merge_range("B7:C7", karyawan[5], idr_format)
            worksheet.merge_range("B8:C8", karyawan[7], idr_format)
            worksheet.merge_range("B9:C9", karyawan[6], idr_format)
            worksheet.merge_range("B10:C10", total_gaji, idr_format)

            workbook.close()

            QMessageBox.about(self, 'Berhasil', 'Slip gaji berhasil diexport')
        else:
            QMessageBox.about(self, 'Peringatan', 'Karyawan tidak ditemukan')

    # Method untuk mereset inputan
    def reset(self):
        self.lineEditNip.setText("")
        self.lineEditNama.setText("")
        self.lineEditGaji.setText("")
        self.lineEditTunjangan.setText("")
        self.lineEditPpn.setText("")
        self.lineEditLembur.setText("")
        self.label_t_gaji.setText("Rp. 0.00")

    # Method untuk menutup widget
    def close_menu(self):
        self.close()

    # Method untuk logout
    def logout(self):
        self.login = LoginWindow()
        self.login.show()
        self.close()
        QMessageBox.about(self, 'Success', 'Logout Berhasil!')


app = QApplication(sys.argv)
window = LoginWindow()
window.show()
app.exec_()
