import sys
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                              QPushButton, QTableView, QFileDialog, QHBoxLayout, 
                              QGridLayout, QTextEdit, QComboBox, QDialog, QHeaderView)
from PySide6.QtCore import Qt, QAbstractTableModel
import pandas as pd
from pathlib import Path
import serial
import serial.tools.list_ports

class SerialPortDialog(QDialog):
    """시리얼 포트 선택을 위한 다이얼로그"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("시리얼 포트 선택")
        self.setModal(True)
        
        layout = QVBoxLayout(self)
        
        # 포트 선택 콤보박스
        self.port_combo = QComboBox()
        self.refresh_ports()
        layout.addWidget(self.port_combo)
        
        # 연결 버튼
        self.connect_button = QPushButton("연결")
        self.connect_button.clicked.connect(self.accept)
        layout.addWidget(self.connect_button)
        
        # 새로고침 버튼
        refresh_button = QPushButton("포트 새로고침")
        refresh_button.clicked.connect(self.refresh_ports)
        layout.addWidget(refresh_button)

    def refresh_ports(self):
        """사용 가능한 시리얼 포트 목록 새로고침"""
        self.port_combo.clear()
        ports = serial.tools.list_ports.comports()
        for port in ports:
            self.port_combo.addItem(f"{port.device} - {port.description}", port.device)

    def get_selected_port(self):
        """선택된 포트 반환"""
        return self.port_combo.currentData()

class TableModel(QAbstractTableModel):
    def __init__(self, data):
        super().__init__()
        self._data = data

    def rowCount(self, index):
        return len(self._data)  # 고정 행 제외

    def columnCount(self, index):
        return len(self._data.columns)

    def data(self, index, role=Qt.DisplayRole):
        if role == Qt.DisplayRole:
            if len(self._data) == 0:
                return None
            value = self._data.iloc[index.row(), index.column()]
            return str(value) if pd.notna(value) and value != '' else '0'
        return None

    def headerData(self, section, orientation, role):
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                return str(self._data.columns[section])
            if orientation == Qt.Vertical:
                return str(section + 1)
        return None

    def addRow(self):
        row_count = len(self._data)
        # 기본값 또는 빈 값으로 새 행 추가
        new_row = pd.Series({col: '' for col in self._data.columns}, index=self._data.columns)
        self._data.loc[row_count] = new_row
        self.layoutChanged.emit()

    def setData(self, index, value, role):
        if role == Qt.EditRole:
            self._data.iloc[index.row(), index.column()] = value
            self.dataChanged.emit(index, index)
            return True
        return False

    def flags(self, index):
        return Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable

class SaveNumberDialog(QDialog):
    """저장 번호 선택을 위한 다이얼로그"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("저장 번호 선택")
        self.setModal(True)
        
        layout = QVBoxLayout(self)
        
        # 번호 선택 콤보박스
        self.number_combo = QComboBox()
        for i in range(1, 5):  # 1~4까지의 번호 추가
            self.number_combo.addItem(f"저장 위치 {i}", i)
        layout.addWidget(self.number_combo)
        
        # 저장 버튼
        self.save_button = QPushButton("저장")
        self.save_button.clicked.connect(self.accept)
        layout.addWidget(self.save_button)

    def get_selected_number(self):
        """선택된 번호 반환"""
        return self.number_combo.currentData()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("패턴 편집기")
        self.setGeometry(100, 100, 1600, 600)

        # 메인 위젯 설정
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout(main_widget)

        # 테이블 뷰 영역 (왼쪽에 위치, 가장 큰 공간 차지)
        center_widget = QWidget()
        center_layout = QVBoxLayout(center_widget)
        
        self.table_view = QTableView()
        self.table_view.setSelectionBehavior(QTableView.SelectRows)
        self.table_view.setSelectionMode(QTableView.SingleSelection)
        
        # 컬럼 너비 설정 (이전 값의 90%로 추가 감소)
        self.table_view.horizontalHeader().setDefaultSectionSize(65)  # 72 * 0.9 = 65
        
        center_layout.addWidget(self.table_view)
        
        main_layout.addWidget(center_widget, stretch=6)

        # 오른쪽 영역 (상태 표시 + 버튼)
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        
        # 상태 표시 영역
        self.status_text = QTextEdit()
        self.status_text.setReadOnly(True)
        self.status_text.setMinimumWidth(250)
        right_layout.addWidget(self.status_text)
        
        # 버튼 영역
        button_widget = QWidget()
        button_layout = QGridLayout(button_widget)
        
        self.add_button = QPushButton("행 추가")
        self.select_button = QPushButton("파일 선택")
        self.delete_button = QPushButton("행 제거")
        self.save_button = QPushButton("저장")
        self.save_as_button = QPushButton("다른 이름으로 저장")
        self.connect_board_button = QPushButton("보드연결")
        self.save_to_board_button = QPushButton("보드저장")
        self.run_board_button = QPushButton("보드구동")
        
        # 버튼을 2열로 배치
        button_layout.addWidget(self.add_button, 0, 0)
        button_layout.addWidget(self.select_button, 0, 1)
        button_layout.addWidget(self.delete_button, 1, 0)
        button_layout.addWidget(self.save_button, 1, 1)
        button_layout.addWidget(self.save_as_button, 2, 0)
        button_layout.addWidget(self.connect_board_button, 2, 1)
        button_layout.addWidget(self.save_to_board_button, 3, 0)
        button_layout.addWidget(self.run_board_button, 3, 1)
        
        right_layout.addWidget(button_widget)
        
        # 오른쪽 영역의 너비 설정
        right_widget.setFixedWidth(300)
        main_layout.addWidget(right_widget, stretch=1)

        # 컬럼 정의 및 나머지 초기화 코드
        self.columns = ['CH1', 'CH2', 'CH3', 'CH4', 'CH5', 'CH6', 'CH7', 'CH8', 'CH9', 
                       'CH10', 'CH11', 'CH12', 'CH13', 'CH14', 'CH15', 'CH16', 'CH17', 'CH18', 'TIME']
        
        self.current_file = Path(__file__).parent / 'data' / 'sequence.csv'
        
        try:
            self.df = pd.read_csv(
                self.current_file,
                names=self.columns,
                header=None
            )
            self.df.dropna(how='all', inplace=True)
            self.update_status("파일을 성공적으로 불러왔습니다.")
        except FileNotFoundError:
            self.df = pd.DataFrame(columns=self.columns)
            self.update_status("새로운 데이터프레임이 생성되었습니다.")

        self.model = TableModel(self.df)
        self.table_view.setModel(self.model)

        # 버튼 이벤트 연결
        self.add_button.clicked.connect(self.add_row)
        self.select_button.clicked.connect(self.select_file)
        self.save_button.clicked.connect(self.save_data)
        self.save_as_button.clicked.connect(self.save_as_data)
        self.delete_button.clicked.connect(self.remove_row)
        self.connect_board_button.clicked.connect(self.connect_board)
        self.save_to_board_button.clicked.connect(self.save_to_board)
        self.run_board_button.clicked.connect(self.run_board)

        # 시리얼 통신 관련 변수 초기화
        self.serial_port = None

    def update_status(self, message):
        """상태 메시지를 업데이트하는 메서드"""
        current_text = self.status_text.toPlainText()
        timestamp = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
        new_message = f"[{timestamp}] {message}\n"
        self.status_text.setText(new_message + current_text)

    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "CSV 파일 선택", "", "CSV Files (*.csv)")
        if file_path:
            self.current_file = Path(file_path)
            try:
                self.df = pd.read_csv(
                    self.current_file,
                    names=self.columns,
                    header=None
                )
                self.df.dropna(how='all', inplace=True)
                self.update_status(f"파일을 불러왔습니다: {file_path}")
            except FileNotFoundError:
                self.df = pd.DataFrame(columns=self.columns)
                self.update_status("파일을 찾을 수 없습니다.")
            
            self.model = TableModel(self.df)
            self.table_view.setModel(self.model)

    def add_row(self):
        self.model.addRow()
        self.update_status("새로운 행이 추가되었습니다.")

    def save_data(self):
        try:
            save_path = self.current_file
            data_to_save = self.df.copy()
            data_to_save = data_to_save.replace('', '0')
            data_to_save.to_csv(save_path, index=False, header=False)
            self.update_status(f"파일이 저장되었습니다: {save_path}")
        except Exception as e:
            self.update_status(f"저장 중 오류 발생: {str(e)}")

    def save_as_data(self):
        """다른 이름으로 저장하는 메서드"""
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "다른 이름으로 저장",
            "",
            "CSV Files (*.csv)"
        )
        if file_path:
            try:
                save_path = Path(file_path)
                data_to_save = self.df.copy()
                data_to_save = data_to_save.replace('', '0')
                data_to_save.to_csv(save_path, index=False, header=False)
                self.current_file = save_path  # 현재 작업 파일을 새로 저장한 파일로 업데이트
                self.update_status(f"파일이 새로운 이름으로 저장되었습니다: {save_path}")
            except Exception as e:
                self.update_status(f"저장 중 오류 발생: {str(e)}")

    def remove_row(self):
        selected_indexes = self.table_view.selectionModel().selectedRows()
        if not selected_indexes:
            self.update_status("선택된 행이 없습니다.")
            return
        for index in sorted(selected_indexes, reverse=True):
            row = index.row()
            self.df = self.df.drop(self.df.index[row])
        self.df.reset_index(drop=True, inplace=True)
        self.model = TableModel(self.df)
        self.table_view.setModel(self.model)
        self.update_status("선택된 행이 삭제되었습니다.")

    def connect_board(self):
        """보드 연결 다이얼로그를 표시하고 연결을 시도"""
        button_text = self.connect_board_button.text()
        if button_text == "보드연결":
            dialog = SerialPortDialog(self)
            if dialog.exec_():
                selected_port = dialog.get_selected_port()
                if selected_port:
                    try:
                        # 이미 연결된 포트가 있다면 닫기
                        if self.serial_port and self.serial_port.is_open:
                            self.serial_port.close()
                            self.update_status("이전 연결이 종료되었습니다.")
                        
                        # 새로운 시리얼 포트 연결
                        self.serial_port = serial.Serial(
                            port=selected_port,
                            baudrate=115200,  # 보드의 통신 속도에 맞게 설정
                            timeout=1
                        )
                        self.update_status(f"보드가 연결되었습니다: {selected_port}")
                        self.connect_board_button.setText("연결 해제")
                    except serial.SerialException as e:
                        self.update_status(f"연결 오류: {str(e)}")
                        self.serial_port = None
                else:
                    self.update_status("포트가 선택되지 않았습니다.")
        else:
            self.serial_port.close()
            self.connect_board_button.setText("보드연결")
            self.update_status(f"보드 연결이 해제되었습니다")

    def closeEvent(self, event):
        """프로그램 종료 시 시리얼 포트 정리"""
        if self.serial_port and self.serial_port.is_open:
            self.serial_port.close()
        event.accept()

    def send_uart_data(self, data):
        """UART를 통해 데이터 전송"""
        if self.serial_port and self.serial_port.is_open:
            try:
                self.serial_port.write(data.encode())
                self.update_status(f"데이터 전송: {data}")
            except serial.SerialException as e:
                self.update_status(f"전송 오류: {str(e)}")
        else:
            self.update_status("보드가 연결되어 있지 않습니다.")

    def read_uart_data(self):
        """UART로부터 데이터 수신"""
        if self.serial_port and self.serial_port.is_open:
            try:
                if self.serial_port.in_waiting:
                    data = self.serial_port.readline().decode().strip()
                    self.update_status(f"수신 데이터: {data}")
                    return data
            except serial.SerialException as e:
                self.update_status(f"수신 오류: {str(e)}")
        return None

    def save_to_board(self):
        """보드에 데이터 저장"""
        if not self.serial_port or not self.serial_port.is_open:
            self.update_status("보드가 연결되어 있지 않습니다.")
            return
            
        dialog = SaveNumberDialog(self)
        if dialog.exec_():
            selected_number = dialog.get_selected_number()
            try:
                # 현재 데이터를 파일에 저장
                self.save_data()
                
                # 데이터프레임의 각 행을 튜플로 변환하여 리스트 생성
                data_list = [(selected_number,)]  # 첫 번째 튜플에 저장 위치 번호 저장
                
                for index, row in self.df.iterrows():
                    # 빈 문자열을 0으로 변환하고 모든 값을 int로 변환
                    row_data = [int(val) if val != '' else 0 for val in row.values]
                    data_list.append(tuple(row_data))
                
                self.update_status(f"저장 위치 {selected_number}번에 데이터를 저장합니다.")
                self.update_status(f"총 {len(data_list)-1}개의 행이 처리되었습니다.")
                self.update_status(f"첫 번째 데이터: {data_list[0]}")
                self.update_status(f"두 번째 데이터: {data_list[1] if len(data_list) > 1 else '없음'}")
                
                # 여기에 실제 보드 저장 로직 구현
                # 예시: self.send_uart_data(f"SAVE,{len(data_list)}")
                # 예시: for data in data_list:
                #          self.send_uart_data(','.join(map(str, data)))
                
            except Exception as e:
                self.update_status(f"보드 저장 중 오류 발생: {str(e)}")

    def run_board(self):
        """보드 구동"""
        if not self.serial_port or not self.serial_port.is_open:
            self.update_status("보드가 연결되어 있지 않습니다.")
            return
            
        try:
            # 현재 데이터를 파일에 저장
            self.save_data()
            
            # 데이터프레임의 각 행을 튜플로 변환하여 리스트 생성
            data_list = []
            for index, row in self.df.iterrows():
                # 빈 문자열을 0으로 변환하고 모든 값을 float로 변환
                row_data = [int(val) if val != '' else 0 for val in row.values]
                data_list.append(tuple(row_data))
            
            self.update_status(f"총 {len(data_list)}개의 행이 처리되었습니다.")
            self.update_status(f"첫 번째 행 데이터: {data_list[0] if data_list else '없음'}")
            
            # 여기에 실제 보드 구동 로직 구현
            # 예시: self.send_uart_data(f"RUN,{len(data_list)}")
            # 예시: for data in data_list:
            #          self.send_uart_data(','.join(map(str, data)))
            
        except Exception as e:
            self.update_status(f"보드 구동 중 오류 발생: {str(e)}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec()) 