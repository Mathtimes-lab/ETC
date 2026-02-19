import sys
from PyQt5.QtWidgets import QApplication
from kiwoom import Kiwoom

if __name__ == "__main__":
    # QApplication 생성
    app = QApplication(sys.argv)

    # Kiwoom 클래스 인스턴스 생성 (강의 이미지 구조 반영)
    kiwoom_window = Kiwoom()

    # 윈도우 창 표시 (강의에서 QMainWindow를 썼으므로 보여주는게 정석입니다)
    # kiwoom_window.show() # 창을 보고 싶다면 주석을 해제하세요.

    # 로그인 실행
    kiwoom_window.comm_connect()

    # 이벤트 루프 실행
    sys.exit(app.exec_())
