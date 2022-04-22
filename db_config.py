import mysql.connector

class Database:
    def __init__(self):

        # Menginisialisasikan attribute dengan konfigurasi ke database
        self._db = mysql.connector.connect(
            host="localhost",
            user="root",
            passwd="",
            database="db_gaji"
        )
        self._cursor = self._db.cursor()


class Karyawan(Database):
    # Menginisialisasikan attribut yang ada di class Database ke dalam self
    def __init__(self):
        Database.__init__(self)

    # Method untuk mengambil semua data karyawan yang ada di database
    def get_karyawan(self, mode="DESC"):
        sql = "SELECT nip, nama, jabatan, gaji, tunjangan, ppn, uang_lembur FROM karyawan ORDER BY id {}".format(mode)

        try:
            self._cursor.execute(sql)
            result = self._cursor.fetchall()
        except Exception as e:
            return e

        # Mengembalikan data result yang berisi list data karyawan
        return result

    # Method untuk mengambil satu data karyawan dari database
    def find_karyawan(self, nip):
        sql = "SELECT * FROM karyawan WHERE nip = {}".format(nip)

        try:
            self._cursor.execute(sql)
            result = self._cursor.fetchone()
        except Exception as e:
            return e

        # Mengembalikan data result dalam bentuk tuple yang berisi data karyawan
        return result

    # Method untuk memasukkan data karyawan ke database
    def store(self, data):
        sql = "INSERT INTO karyawan (nip, nama, jabatan,uang_lembur, gaji, tunjangan, ppn)" \
              " VALUES (%s, %s, %s, %s, %s, %s, %s) "

        try:
            self._cursor.execute(sql, data)
            self._db.commit()
        except Exception as e:
            return e

        # Mengembalikan nilai berupa id terakhir karyawan
        return self._cursor.lastrowid

    # Method untuk mengubah data karyawan ke database berdasarkan id
    def update(self, id, data):
        sql = "UPDATE karyawan SET nip = %s,nama = %s,jabatan = %s,uang_lembur = %s,gaji = %s,tunjangan = %s,ppn = %s WHERE id = {}".format(id)

        try:
            self._cursor.execute(sql, data)
            self._db.commit()
        except Exception as e:
            return e

    # Method untuk menghapus data karyawan berdasarkan id
    def delete(self, id):
        sql = "DELETE FROM karyawan WHERE id = {}".format(id)

        try:
            self._cursor.execute(sql)
            self._db.commit()
        except Exception as e:
            return e
