import sys
from PyQt5 import QtCore, QtGui, QtWidgets
from googleapiclient.discovery import build
from google.oauth2 import service_account
from datetime import datetime
import pygame

app = QtWidgets.QApplication(sys.argv)
window = QtWidgets.QMainWindow()
window.setWindowTitle('Ni Fashion Software')
window.resize(800, 600)
central_widget = QtWidgets.QWidget(window)
window.setCentralWidget(central_widget)
layout = QtWidgets.QVBoxLayout(central_widget)
layout.setAlignment(QtCore.Qt.AlignCenter)

frame = QtWidgets.QFrame()
frame.setStyleSheet("background-color: white; border-radius: 20px;")
frame.setFixedSize(600, 500)
frame_layout = QtWidgets.QVBoxLayout(frame)

lbl_awb = QtWidgets.QLabel('Enter AWB')
lbl_awb.setFont(QtGui.QFont('Courier New', 14))
lbl_awb.setStyleSheet('background: white; color: black; font-weight: bold; border-radius: 10px; padding: 5px;')

entry_awb = QtWidgets.QLineEdit()
entry_awb.setStyleSheet("QLineEdit {background-color: white; border: 1px solid gray; border-radius: 5px; padding: 5px; font-size: 14px;} QLineEdit:hover {background-color: lightgray; border-color: darkgray;}")

awb_layout = QtWidgets.QHBoxLayout()
awb_layout.addWidget(lbl_awb)
awb_layout.addWidget(entry_awb)

lbl_validation = QtWidgets.QLabel('')
lbl_validation.setFont(QtGui.QFont('Arial', 19))
lbl_validation.setStyleSheet("QLabel {border-radius: 5px; padding: 15px; background-color: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #4169E1, stop:1 #8A2BE2); color: white; font-weight: bold; qproperty-alignment: AlignCenter; qproperty-wordWrap: true; background-clip: content; border-image: none; margin: 10px;} QLabel:hover {background-color: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #483D8B, stop:1 #8A2BE2); border-color: #483D8B;}")

report_treeview = QtWidgets.QTreeView()
report_treeview.setRootIsDecorated(False)
report_treeview.setAlternatingRowColors(True)
report_treeview.setFixedHeight(200)
report_treeview.setStyleSheet("QTreeView::item { border: 1px solid black; }")

report_model = QtGui.QStandardItemModel()
report_model.setColumnCount(7)
report_model.setHorizontalHeaderLabels(['GOOGLE SHEETS', '|||||||||||', '|||||||||||', '|||||||||||', '|||||||||||', '|||||||||||', '(!) Pending'])
report_treeview.setModel(report_model)

header = report_treeview.header()
header.setSectionResizeMode(QtWidgets.QHeaderView.ResizeToContents)
header.setStyleSheet("QHeaderView::section {background-color: #44A44D; border: none; font-weight: bold; padding: 5px; color: white;}")

# Set the style sheet for the Treeview
report_treeview.setStyleSheet("QTreeView::item { border: 1px solid black; } QTreeView::item:selected { background-color: blue; color: white; }")

treeview_scrollbar = QtWidgets.QScrollBar(QtCore.Qt.Vertical)
report_treeview.setVerticalScrollBar(treeview_scrollbar)
treeview_scrollbar.setMaximum(report_treeview.height())

pygame.mixer.init()
sound_files = {'no_data': 'no_data.wav', 'success': 'success.wav', 'invalid': 'invalid.wav', 'duplicate': 'duplicate.wav'}
image_files = {'no_data': 'no_data_image.png', 'success': 'success_image.png', 'invalid': 'invalid_image.png', 'duplicate': 'duplicate_image.png'}

def play_sound(outcome):
    sound_file = sound_files.get(outcome)
    if sound_file:
        pygame.mixer.music.load(sound_file)
        pygame.mixer.music.play()

def display_image(outcome):
    image_file = image_files.get(outcome)
    if image_file:
        image_label.setPixmap(QtGui.QPixmap(image_file))

def submit_data():
    awb = entry_awb.text()
    if not awb:
        lbl_validation.setText('(!) ERROR: No data entered!')
        play_sound('no_data')
        display_image('no_data')
        return

    creds = service_account.Credentials.from_service_account_file('scan-390611-25fc1daa0744.json', scopes=['https://www.googleapis.com/auth/spreadsheets'])
    service = build('sheets', 'v4', credentials=creds)

    validation_range = 'Manifest!D:F'
    result = service.spreadsheets().values().get(spreadsheetId='1JvvmA_Lbr55XwYD2atYzeoAbD_ntOJxPx9IDKyCbK5c', range=validation_range).execute()
    validation_data = result.get('values', [])

    valid, company_name, tpl_name = False, '', ''
    for row in validation_data:
        if len(row) > 2 and awb == row[0]:
            valid = True
            company_name, tpl_name = row[1], row[2]
            lbl_validation.setText(f'Valid AWB! Company: {company_name} 3PL: {tpl_name}')
            break

    if valid:
        data_range = 'Data!A:A'
        result = service.spreadsheets().values().get(spreadsheetId='1JvvmA_Lbr55XwYD2atYzeoAbD_ntOJxPx9IDKyCbK5c', range=data_range).execute()
        data_values = result.get('values', [])
        if any(awb in row for row in data_values):
            lbl_validation.setText(f"(!) That's a DUPLICATE, already scanned '{awb}'")
            play_sound('duplicate')
            display_image('duplicate')
        else:
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            values = [awb, now]
            body = {'values': [values]}
            service.spreadsheets().values().append(spreadsheetId='1JvvmA_Lbr55XwYD2atYzeoAbD_ntOJxPx9IDKyCbK5c', range='Data!A:B', valueInputOption='USER_ENTERED', insertDataOption='INSERT_ROWS', body=body).execute()
            lbl_validation.setText(f"(✓) Scanned '{awb}' of '{company_name}' will be delivered by '{tpl_name}'")
            play_sound('success')
            display_image('success')
    else:
        lbl_validation.setText("(!) Sorry, Order CANCELLED or invalid AWB")
        play_sound('invalid')
        display_image('invalid')

    entry_awb.clear()

    result = service.spreadsheets().values().get(spreadsheetId='1JvvmA_Lbr55XwYD2atYzeoAbD_ntOJxPx9IDKyCbK5c', range='Report!Report1').execute()
    report_data = result.get('values', [])

    report_model.removeRows(0, report_model.rowCount())

    for row in report_data:
        report_model.appendRow([QtGui.QStandardItem(data)for data in row])

image_label = QtWidgets.QLabel()

frame_layout.addLayout(awb_layout)
frame_layout.addWidget(report_treeview)
frame_layout.addWidget(lbl_validation)
frame_layout.addWidget(image_label)

entry_awb.returnPressed.connect(submit_data)

layout.addWidget(frame)

window.show()
sys.exit(app.exec_())
