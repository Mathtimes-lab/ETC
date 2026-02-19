import sys
import time
from datetime import datetime
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
from PyQt5.QtCore import *


class Kiwoom(QMainWindow):
    def __init__(self):
        super().__init__()

        # --- [설정 구간] ---
        self.BUY_STRATEGY_NAME = "260218급등기본"
        self.SELL_STRATEGY_NAME = "260218매도식"
        # -----------------

        self.account_num = None
        self.bought_today = []  # 오늘 매수한 종목 리스트
        self.held_stocks = []  # 현재 보유중인 종목 리스트
        self.open_buy_orders = {}  # 미체결 매수 주문 관리

        # 키움 OCX 생성
        self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")

        # API 팝업 억제
        self.kiwoom.dynamicCall("KOA_Functions(QString, QString)", "SetShowMessage", "0")

        # 이벤트 연결
        self.kiwoom.OnEventConnect.connect(self._event_connect)
        self.kiwoom.OnReceiveConditionVer.connect(self._handler_condition_load)
        self.kiwoom.OnReceiveTrData.connect(self._handler_tr_data)
        self.kiwoom.OnReceiveMsg.connect(self._handler_msg)  # [추가] 서버 메시지(주문실패 사유 등) 수신

        try:
            self.kiwoom.OnReceiveTrCondition.connect(self._handler_condition)
        except AttributeError:
            self.kiwoom.OnReceiveCondition.connect(self._handler_condition)

        self.kiwoom.OnReceiveRealCondition.connect(self._handler_real_condition)
        self.kiwoom.OnReceiveChejanData.connect(self._handler_chejan_data)

        # 미체결 주문 취소 타이머 (1분)
        self.cancel_timer = QTimer(self)
        self.cancel_timer.timeout.connect(self._check_time_and_cancel)
        self.cancel_timer.start(60000)

    # -------------------------------------
    # [신규] 호가 단위 계산 함수 (KOSPI/KOSDAQ 통합 보정)
    # -------------------------------------
    def _get_hoga_unit(self, price):
        """가격대별 호가 단위 반환"""
        if price < 2000:
            return 1
        elif price < 5000:
            return 5
        elif price < 20000:
            return 10
        elif price < 50000:
            return 50
        elif price < 200000:
            return 100
        elif price < 500000:
            return 500
        else:
            return 1000

    def _adjust_price_to_tick(self, price):
        """계산된 가격을 호가 단위에 맞게 보정"""
        unit = self._get_hoga_unit(price)
        # 호가 단위로 반올림 처리
        adjusted_price = int(round(price / unit) * unit)
        return adjusted_price

    # -------------------------------------
    # 초기화 체인
    # -------------------------------------
    def _req_outstanding_orders(self):
        """미체결 내역 요청"""
        print("[시스템] 미체결 주문 내역을 확인합니다...")
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "전체종목구분", "0")
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "매매구분", "0")
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "체결구분", "1")
        self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "미체결요청", "opt10075", 0, "0102")

    def _req_account_balance(self):
        """보유 종목(잔고) 요청"""
        print("[시스템] 보유 종목(계좌 잔고)을 확인합니다...")
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "비밀번호", "")
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "조회구분", "2")
        self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "잔고요청", "opw00018", 0, "0103")

    def _handler_tr_data(self, scr_no, rqname, trcode, recordname, prev_next, data_len, err_code, msg, splm_msg):
        if rqname == "미체결요청":
            cnt = self.kiwoom.dynamicCall("GetRepeatCnt(QString, QString)", trcode, rqname)
            for i in range(cnt):
                code = self.kiwoom.dynamicCall("GetCommData(QString, QString, int, QString)", trcode, rqname, i,
                                               "종목코드").strip()
                order_no = self.kiwoom.dynamicCall("GetCommData(QString, QString, int, QString)", trcode, rqname, i,
                                                   "주문번호").strip()
                order_type = self.kiwoom.dynamicCall("GetCommData(QString, QString, int, QString)", trcode, rqname, i,
                                                     "주문구분").strip()

                if "매수" in order_type:
                    self.open_buy_orders[code] = order_no

            print(f"[시스템] 미체결 매수 주문 복원: {len(self.open_buy_orders)}건")
            QTimer.singleShot(200, self._req_account_balance)

        elif rqname == "잔고요청":
            cnt = self.kiwoom.dynamicCall("GetRepeatCnt(QString, QString)", trcode, rqname)
            for i in range(cnt):
                code = self.kiwoom.dynamicCall("GetCommData(QString, QString, int, QString)", trcode, rqname, i,
                                               "종목번호").strip()
                code = code.replace("A", "")
                if code not in self.held_stocks:
                    self.held_stocks.append(code)

            print(f"[시스템] 보유 종목 리스트 복원: {len(self.held_stocks)}종목")
            QTimer.singleShot(200, self._get_condition_load)

    # -------------------------------------
    # 메인 로직
    # -------------------------------------
    def _execute_buy(self, code):
        """매수 로직: 전일종가 기준 +5% 가격 계산(호가보정) -> 주문"""
        now = datetime.now().strftime('%H:%M:%S')
        stock_name = self.kiwoom.dynamicCall("GetMasterCodeName(QString)", code)

        if code in self.open_buy_orders:
            print(f"[{now}] [매수스킵] {stock_name}({code}) - 미체결 매수 주문 대기 중")
            return

        if code in self.held_stocks:
            print(f"[{now}] [매수스킵] {stock_name}({code}) - 이미 보유 중인 종목")
            return

        if code in self.bought_today:
            return

        prev_price_str = self.kiwoom.dynamicCall("GetMasterLastPrice(QString)", code)
        prev_price = abs(int(prev_price_str)) if prev_price_str else 0

        if prev_price == 0: return

        # 1. 목표 매수 단가 계산 (전일 종가 + 5%)
        raw_target_price = prev_price * 1.05

        # 2. [핵심] 호가 단위 보정 (120,015원 -> 120,000원 등으로 보정)
        target_price = self._adjust_price_to_tick(raw_target_price)

        # 3. 수량 계산
        quantity = 1000000 // target_price
        if quantity == 0:
            print(f"[{now}] [매수불가] {stock_name}({code}) - 단가가 100만원 초과")
            return

        print(f"[{now}] [자동매수] {stock_name}({code})")
        print(f"  > 전일종가: {prev_price:,}원")
        print(f"  > 주문가격: {target_price:,}원 (호가보정됨)")
        print(f"  > 주문수량: {quantity}주")

        self._send_order(code, 1, quantity, target_price)

        if code not in self.bought_today:
            self.bought_today.append(code)

    def _execute_sell(self, code):
        """매도 로직"""
        now = datetime.now().strftime('%H:%M:%S')
        stock_name = self.kiwoom.dynamicCall("GetMasterCodeName(QString)", code)

        if code in self.bought_today:
            print(f"[{now}] [매도스킵] {stock_name}({code}) - 당일 매수 종목")
            return

        print(f"[{now}] [자동매도] {stock_name}({code}) 시장가 주문 전송")
        self._send_order(code, 2, 10, 0)

    # -------------------------------------
    # 이벤트 핸들러
    # -------------------------------------
    def comm_connect(self):
        print("[시스템] 로그인 창을 불러옵니다...")
        self.kiwoom.dynamicCall("CommConnect()")
        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_()

    def _event_connect(self, err_code):
        if err_code == 0:
            print("로그인 성공!")
            self._get_account_info()
            QTimer.singleShot(200, self._req_outstanding_orders)
        else:
            print("로그인 실패!")
            if hasattr(self, 'login_event_loop'): self.login_event_loop.exit()

    def _handler_msg(self, scr_no, rqname, trcode, msg):
        """[추가] 서버 메시지 수신 (주문 실패 사유 확인용)"""
        # "매수" 관련 로그만 필터링해서 보여줍니다.
        if "매수" in rqname or "주문" in msg:
            print(f"[서버메시지] {msg}")

    def after_login(self):
        if hasattr(self, 'login_event_loop'): self.login_event_loop.exit()

    def _get_account_info(self):
        account_list = self.kiwoom.dynamicCall("GetLoginInfo(QString)", "ACCNO")
        self.account_num = account_list.split(';')[0]
        print(f"[내 정보] 계좌번호: {self.account_num}")

    def _get_condition_load(self):
        print("[시스템] 조건식 목록 요청 중...")
        self.kiwoom.dynamicCall("GetConditionLoad()")

    def _handler_condition_load(self, ret, msg):
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
                    print(f"[{name}] 초기 검색 요청...")
                    self.kiwoom.dynamicCall("SendCondition(QString, QString, int, int)", "0157", name, int(index), 0)

    def _handler_condition(self, scr_no, code_list, cond_name, cond_index, next):
        now = datetime.now().strftime('%H:%M:%S')

        if not code_list or code_list.strip() == "":
            print(f"[{now}] [{cond_name}] 검색된 종목 없음")
        else:
            codes = code_list.split(';')[:-1]
            print("\n" + "=" * 50)
            print(f"[{now}] [{cond_name}] 초기 검색 결과: {len(codes)}종목 발견 -> 자동 주문 진행")
            print("-" * 50)

            for code in codes:
                if cond_name == self.BUY_STRATEGY_NAME:
                    self._execute_buy(code)
                elif cond_name == self.SELL_STRATEGY_NAME:
                    self._execute_sell(code)
                # 주문 간격 0.3초로 약간 늘림 (안정성 확보)
                time.sleep(0.3)

            print("=" * 50 + "\n")

        print(f"[시스템] {cond_name} 실시간 감시 전환 (화면번호: {scr_no})")
        self.kiwoom.dynamicCall("SendCondition(QString, QString, int, int)", scr_no, cond_name, int(cond_index), 1)

    def _handler_real_condition(self, code, type, cond_name, cond_index):
        if type == 'I':  # 종목 편입
            if cond_name == self.BUY_STRATEGY_NAME:
                self._execute_buy(code)
            elif cond_name == self.SELL_STRATEGY_NAME:
                self._execute_sell(code)

    def _send_order(self, code, order_type, quantity, price):
        order_name = "매수" if order_type == 1 else "매도"
        if order_type == 3: order_name = "취소"

        hoga_gb = "00"
        if order_type == 2:
            hoga_gb = "03"
            price = 0

        ret = self.kiwoom.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                                      ["send_order", "0101", self.account_num, order_type, code, quantity, price,
                                       hoga_gb, ""])
        if ret == 0:
            print(f"  ==> {order_name} 주문 전송: {code}")
        else:
            print(f"  ==> {order_name} 실패: {code} (에러: {ret})")

    def _handler_chejan_data(self, gubun, item_cnt, fid_list):
        if gubun == '0':  # 주문/체결
            status = self.kiwoom.dynamicCall("GetChejanData(int)", 913)
            code = self.kiwoom.dynamicCall("GetChejanData(int)", 9001).replace('A', '')
            order_no = self.kiwoom.dynamicCall("GetChejanData(int)", 9203)
            order_type = self.kiwoom.dynamicCall("GetChejanData(int)", 905)

            if "매수" in order_type and status == "접수":
                self.open_buy_orders[code] = order_no
            elif "매수" in order_type and status == "체결":
                if code in self.open_buy_orders:
                    del self.open_buy_orders[code]
                if code not in self.held_stocks:
                    self.held_stocks.append(code)

            print(f"[체결알림] {code} | {status} | 주문번호: {order_no}")

        elif gubun == '1':  # 잔고통보
            code = self.kiwoom.dynamicCall("GetChejanData(int)", 9001).replace('A', '')
            if code not in self.held_stocks:
                self.held_stocks.append(code)

    def _check_time_and_cancel(self):
        now = datetime.now()
        if now.hour == 15 and now.minute >= 20:
            if self.open_buy_orders:
                print(f"[시스템] 장 마감 임박(15:20) - 미체결 매수 주문을 일괄 취소합니다.")
                for code, order_no in list(self.open_buy_orders.items()):
                    self.kiwoom.dynamicCall(
                        "SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                        ["cancel_order", "0102", self.account_num, 3, code, 0, 0, "00", order_no])
                    del self.open_buy_orders[code]
                    time.sleep(0.2)
