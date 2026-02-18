import sys
from datetime import datetime
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
from PyQt5.QtCore import *


class Kiwoom(QMainWindow):
    def __init__(self):
        super().__init__()

        # --- [설정 구간] HTS에서 저장한 조건식 이름을 정확히 입력하세요 ---
        self.BUY_STRATEGY_NAME = "260218급등기본"
        self.SELL_STRATEGY_NAME = "260218매도식"
        # -----------------------------------------------------------

        self.account_num = None
        self.bought_today = []  # 오늘 매수한 종목 리스트 (당일 매도 방지용)
        self.open_buy_orders = {}  # 미체결 매수 주문 관리 {종목코드: 주문번호}

        # 키움 OCX 생성
        self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")

        # [추가] API 자체 팝업 메시지 박스 억제
        # 이 함수를 호출하면 API 내부적으로 발생하는 알림 창들을 숨길 수 있습니다.
        self.kiwoom.dynamicCall("KOA_Functions(QString, QString)", "SetShowMessage", "0")

        # 이벤트 연결
        self.kiwoom.OnEventConnect.connect(self._event_connect)
        self.kiwoom.OnReceiveConditionVer.connect(self._handler_condition_load)

        try:
            self.kiwoom.OnReceiveTrCondition.connect(self._handler_condition)
        except AttributeError:
            self.kiwoom.OnReceiveCondition.connect(self._handler_condition)

        self.kiwoom.OnReceiveRealCondition.connect(self._handler_real_condition)
        self.kiwoom.OnReceiveChejanData.connect(self._handler_chejan_data)

        # 미체결 주문 취소를 위한 타이머 설정 (1분마다 체크)
        self.cancel_timer = QTimer(self)
        self.cancel_timer.timeout.connect(self._check_time_and_cancel)
        self.cancel_timer.start(60000)  # 60초

    def comm_connect(self):
        """로그인 실행"""
        print("[시스템] 로그인 창을 불러옵니다...")
        self.kiwoom.dynamicCall("CommConnect()")
        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_()

    def _event_connect(self, err_code):
        """로그인 결과 처리"""
        if err_code == 0:
            print("로그인 성공!")
            self._get_account_info()
            self._get_condition_load()
        else:
            print("로그인 실패!")

        if hasattr(self, 'login_event_loop'):
            self.login_event_loop.exit()
        self.after_login()

    def after_login(self):
        """연결 상태 확인"""
        if self.kiwoom.dynamicCall("GetConnectState()") == 0:
            print("서버와 연결이 끊겼습니다!")
        else:
            print("서버와 연결 중입니다!")

    def _get_account_info(self):
        """계좌 정보 가져오기"""
        account_list = self.kiwoom.dynamicCall("GetLoginInfo(QString)", "ACCNO")
        self.account_num = account_list.split(';')[0]
        print(f"[내 정보] 계좌번호: {self.account_num}")

    def _get_condition_load(self):
        """조건식 목록 요청"""
        print("[시스템] 조건식 목록 요청 중...")
        self.kiwoom.dynamicCall("GetConditionLoad()")

    def _handler_condition_load(self, ret, msg):
        """조건식 로딩 완료 후 초기 리스트 조회 시작"""
        if ret == 1:
            print("[시스템] 조건식 로딩 완료. 초기 종목 스캔을 시작합니다.")
            condition_info = self.kiwoom.dynamicCall("GetConditionNameList()")
            conditions = condition_info.split(";")[:-1]

            for condition in conditions:
                index, name = condition.split('^')
                if name == self.BUY_STRATEGY_NAME:
                    print(f"[{name}] 초기 검색 요청...")
                    self.kiwoom.dynamicCall("SendCondition(QString, QString, int, int)", "0156", name, int(index), 0)
                elif name == self.SELL_STRATEGY_NAME:
                    self.kiwoom.dynamicCall("SendCondition(QString, QString, int, int)", "0157", name, int(index), 0)

    def _handler_condition(self, scr_no, code_list, cond_name, cond_index, next):
        """초기 검색 결과 출력 및 실시간 감시 전환"""
        now = datetime.now().strftime('%H:%M:%S')

        # 검색 결과가 없는 경우 처리 (이미 코드상에서 처리됨)
        if not code_list or code_list.strip() == "":
            print(f"[{now}] [{cond_name}] 현재 검색된 종목이 없습니다.")
        else:
            codes = code_list.split(';')[:-1]
            print("\n" + "=" * 50)
            print(f"[{now}] [{cond_name}] 초기 검색 리스트 (총 {len(codes)}종목)")
            print("-" * 50)
            for code in codes:
                stock_name = self.kiwoom.dynamicCall("GetMasterCodeName(QString)", code)
                prev_price = abs(int(self.kiwoom.dynamicCall("GetMasterLastPrice(QString)", code)))
                print(f"  > 종목: {stock_name}({code}) | 전일가: {prev_price:,}원")
            print("=" * 50 + "\n")

        # 실시간 모니터링으로 자동 전환
        print(f"[시스템] {cond_name} 실시간 모니터링 모드 활성화")
        self.kiwoom.dynamicCall("SendCondition(QString, QString, int, int)", scr_no, cond_name, int(cond_index), 1)

    def _handler_real_condition(self, code, type, cond_name, cond_index):
        """실시간 종목 포착 시 즉시 매매 실행"""
        now_dt = datetime.now()
        now_str = now_dt.strftime('%H:%M:%S')

        if type == 'I':  # 종목 편입
            curr_price_str = self.kiwoom.dynamicCall("GetMasterLastPrice(QString)", code)
            curr_price = abs(int(curr_price_str)) if curr_price_str else 0
            stock_name = self.kiwoom.dynamicCall("GetMasterCodeName(QString)", code)

            if curr_price == 0: return

            if cond_name == self.BUY_STRATEGY_NAME:
                # 100만원 한도 내 최대 수량 계산
                quantity = 1000000 // curr_price
                if quantity == 0: return

                # 즉시 지정가 매수 주문 (팝업 없음)
                print(f"[{now_str}] [자동매수] {stock_name}({code}) | 가격: {curr_price:,}원 | 수량: {quantity}주")
                self._send_order(code, 1, quantity, curr_price)

                if code not in self.bought_today:
                    self.bought_today.append(code)

            elif cond_name == self.SELL_STRATEGY_NAME:
                # 당일 매수 종목 보호
                if code in self.bought_today:
                    print(f"[{now_str}] [매도스킵] {stock_name}({code}) - 당일 매수 종목")
                else:
                    # 즉시 시장가 매도 주문 (팝업 없음)
                    print(f"[{now_str}] [자동매도] {stock_name}({code}) 시장가 주문 전송")
                    self._send_order(code, 2, 10, 0)  # 수량 10주 예시 (잔고조회 연동 권장)

    def _send_order(self, code, order_type, quantity, price):
        """
        주문 전송 함수
        order_type: 1(매수), 2(매도), 3(매수취소)
        매수: 지정가(00), 매도: 시장가(03)
        """
        order_name = ""
        hoga_gb = "00"  # 기본 지정가

        if order_type == 1:
            order_name = "신규매수"
            hoga_gb = "00"  # 지정가 매수
        elif order_type == 2:
            order_name = "신규매도"
            hoga_gb = "03"  # 시장가 매도
            price = 0  # 시장가는 가격 0
        elif order_type == 3:
            order_name = "매수취소"
            hoga_gb = "00"

        # SendOrder 호출
        ret = self.kiwoom.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                                      ["send_order", "0101", self.account_num, order_type, code, quantity, price,
                                       hoga_gb, ""])

        if ret == 0:
            print(f"  ==> {order_name} 요청 성공: {code}")
        else:
            print(f"  ==> {order_name} 요청 실패: {code} (코드: {ret})")

    def _handler_chejan_data(self, gubun, item_cnt, fid_list):
        """체결 데이터 확인 및 미체결 주문 관리"""
        if gubun == '0':  # 주문 접수/체결
            status = self.kiwoom.dynamicCall("GetChejanData(int)", 913)  # 주문상태
            code = self.kiwoom.dynamicCall("GetChejanData(int)", 9001).replace('A', '')  # 종목코드
            order_no = self.kiwoom.dynamicCall("GetChejanData(int)", 9203)  # 주문번호
            order_type = self.kiwoom.dynamicCall("GetChejanData(int)", 905)  # 주문구분 (+매수, -매도)

            # 매수 주문이고 아직 체결 전(접수)인 경우 미체결 리스트에 관리
            if "매수" in order_type:
                if status == "접수":
                    self.open_buy_orders[code] = order_no
                elif status == "체결":
                    # 전량 체결 시 리스트에서 제거 (부분 체결은 유지될 수 있으나 단순화)
                    if code in self.open_buy_orders:
                        del self.open_buy_orders[code]

            print(f"[FID알림] 종목: {code} | 상태: {status} | 주문번호: {order_no}")

    def _check_time_and_cancel(self):
        """장 종료 10분 전(15:20) 미체결 매수 주문 자동 취소"""
        now = datetime.now()
        # 15시 20분 이후이고 미체결 주문이 있다면 취소 실행
        if now.hour == 15 and now.minute >= 20:
            if self.open_buy_orders:
                print(f"[{now.strftime('%H:%M:%S')}] 장 종료 임박 - 미체결 매수 주문을 일괄 취소합니다.")
                # 딕셔너리 복사본으로 반복 (반복 중 삭제 방지)
                for code, order_no in list(self.open_buy_orders.items()):
                    # 취소 주문 전송 (nOrderType: 3)
                    # 취소 시 수량은 0으로 보내면 전량 취소됨
                    self.kiwoom.dynamicCall(
                        "SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                        ["cancel_order", "0102", self.account_num, 3, code, 0, 0, "00", order_no])
                    print(f"  ! 취소 전송: {code} (원주문번호: {order_no})")
                    del self.open_buy_orders[code]
