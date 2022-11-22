import PySimpleGUI as sg
import pandas as pd

sg.theme('DarkBrown')
sg.set_options(font=('montserrat 16'), text_color='white')

EXCEL_FILE = 'Data Warga.xlsx'

df = pd.read_excel(EXCEL_FILE)

layout = [
    [sg.Text('Pencatatan Data Warga')],
    [sg.Text('Nama Kepala Keluarga', size = (21, 1)), sg.InputText(key = 'Nama Kepala Keluarga')],
    [sg.Text('Jenis Kelamin', size = (21, 1)), sg.Combo(['Pria', 'Wanita'], key = 'Jenis Kelamin')],
    [sg.Text('Nomor Kartu Keluarga', size = (21, 1)), sg.InputText(key = 'Nomor Kartu Keluarga')],
    [sg.Text('Alamat', size = (21, 1)), sg.InputText(key = 'Alamat')],
    [sg.Text('Nomor Telpon', size = (21, 1)), sg.InputText(key = 'Nomor Telpon')],
    [sg.Text('Status Perkawinan', size = (21, 1)), sg.Combo(['Kawin', 'Lajang', 'Kawin Cerai'], key = 'Status Perkawinan')],
    [sg.Text('Jumlah Anak', size = (21, 1)), sg.InputText(key = 'Jumlah Anak')],
    [sg.Button('Submit', expand_x=True), sg.Button('Clear', expand_x=True), sg.Button('Exit', expand_x=True)]

]

window = sg.Window('Form Pencatatan Data Warga', layout)

def clear_input():
    for key in values:
        window['Nama Kepala Keluarga'].update('')
        window['Jenis Kelamin'].update('')
        window['Nomor Kartu Keluarga'].update('')
        window['Nomor Telpon'].update('')
        window['Alamat'].update('')
        window['Status Perkawinan'].update('')
        window['Jumlah Anak'].update('')
        return None

while True :
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Clear':
        clear_input()
    if event == 'Submit':
        Name = values['Nama Kepala Keluarga']
        if Name == '':
            sg.PopupError('Masukkan Nama Kepala Keluarga Terlebih Dahulu')
        NoKK = values['Nomor Kartu Keluarga']
        if NoKK == '':
            sg.PopupError('Masukkan Nomor Kartu Keluarga Terlebih Dahulu')
        Jenis = values['Jenis Kelamin']
        if Jenis == '':
            sg.PopupError('Masukkan Jenis Kelamin Terlebih Dahulu')
        NoTLP = values['Nomor Telpon']
        if NoTLP == '':
            sg.PopupError('Masukkan Nomor Telpon Terlebih Dahulu')
        AL = values['Alamat']
        if AL == '':
            sg.PopupError('Masukkan Alamat Terlebih Dahulu')
        SP = values['Status Perkawinan']
        if SP == '':
            sg.PopupError('Masukkan Status Perkawinan Terlebih Dahulu')
        JA = values['Jumlah Anak']
        if JA == '':
            sg.PopupError('Masukkan Jumlah Anak Terlebih Dahulu')
        else:
            try:
                summary_list = 'Informasi Sudah Tersimpan di Database'
                na="\nNama Kepala Keluarga:" + values['Nama Kepala Keluarga']
                na="\nJenis Kelamin:" + values['Jenis Kelamin']
                na="\nNomor Kartu Keluarga:" + values['Nomor Kartu Keluarga']
                na="\nNomor Telpon:" + values['Nomor Telpon']
                na="\nAlamat:" + values['Alamat']
                na="\nStatus Perkawinan:" + values['Status Perkawinan']
                na="\nJumlah Anak:" + values['Jumlah Anak']

                sg.PopupOKCacnel(summary_list, 'Mohon Coba Kembali')
            except:
                choice = sg.PopupError("Error Tidak di Ketahui, Mohon Laporkan Ke Admin")
                if choice == 'OK':
                    clear_input()
                    sg.PopupQuick('Tersimpan')
                else:
                    sg.PopupOK('Edit Ulang')
        if event == 'Submit':
            df = df.append(values, ignore_index=True)
            df.to_excel(EXCEL_FILE, index=False)
            sg.popup('Data Berhasil Tersimpan')
            clear_input()
window.close()       
