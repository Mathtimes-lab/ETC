import sys
import time
import os
import pandas as pd  # 데이터 저장/수정용
import numpy as np  # 날짜 계산용
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
        self.CSV_FILE_NAME = "trade_history.csv"  # 저장할 파일명
        # -----------------

        self.account_num = None
        self.bought_today = []  # 오늘 매수한 종목 리스트

        # 보유 종목 관리: { '종목코드': {'qty': 수량, 'price': 단가, ...} }
        self.held_stocks = {}

        # 매수 주문 시점의 메타 데이터 임시 저장 (체결 전까지 보관)
        self.buy_meta_data = {}

        self.open_buy_orders = {}  # 미체결 매수 주문 관리

        # 현재 실시간 조건 만족 종목 관리 (모니터링용)
        self.current_conditioned_stocks = set()

        # 키움 OCX 생성
        self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")

        # API 팝업 억제
        self.kiwoom.dynamicCall("KOA_Functions(QString, QString)", "SetShowMessage", "0")

        # 이벤트 연결
        self.kiwoom.OnEventConnect.connect(self._event_connect)
        self.kiwoom.OnReceiveConditionVer.connect(self._handler_condition_load)
        self.kiwoom.OnReceiveTrData.connect(self._handler_tr_data)
        self.kiwoom.OnReceiveMsg.connect(self._handler_msg)

        try:
            self.kiwoom.OnReceiveTrCondition.connect(self._handler_condition)
        except AttributeError:
            self.kiwoom.OnReceiveCondition.connect(self._handler_condition)

        self.kiwoom.OnReceiveRealCondition.connect(self._handler_real_condition)
        self.kiwoom.OnReceiveChejanData.connect(self._handler_chejan_data)

        # 시스템 점검 및 미체결 취소 타이머 (1분)
        self.periodic_timer = QTimer(self)
        self.periodic_timer.timeout.connect(self._periodic_check)
        self.periodic_timer.start(60000)

        # 슬리피지 분석 리포트 타이머 (5분)
        self.slippage_timer = QTimer(self)
        self.slippage_timer.timeout.connect(self._print_slippage_report)
        self.slippage_timer.start(300000)

        # -------------------------------------

    # [유틸] 호가 단위 계산
    # -------------------------------------
    def _get_hoga_unit(self, price):
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
        unit = self._get_hoga_unit(price)
        return int(round(price / unit) * unit)

    # -------------------------------------
    # [핵심] CSV 이력 관리 로직 (분리형)
    # -------------------------------------
    def _log_buy_trade(self, code, stock_name, buy_date, buy_time, target_raw, buy_price, slippage):
        """매수 체결 시: 신규 행 생성 (매도 정보는 공란)"""

        # 중복 저장 방지 (이미 같은 종목, 날짜, 시간에 매수 기록이 있으면 스킵)
        if os.path.exists(self.CSV_FILE_NAME):
            try:
                df_exist = pd.read_csv(self.CSV_FILE_NAME, dtype={'종목코드': str})
                # 동일한 매수 건이 있는지 확인
                is_duplicate = not df_exist[
                    (df_exist['종목코드'] == code) &
                    (df_exist['매수일'] == buy_date) &
                    (df_exist['매수시간'] == buy_time)
                    ].empty
                if is_duplicate:
                    return  # 중복이면 저장 안 함
            except Exception:
                pass

        new_data = {
            '종목코드': code,
            '종목명': stock_name,
            '매수일': buy_date,
            '매수시간': buy_time,
            '5%상승가(보정X)': int(target_raw),
            '실제매입가': buy_price,
            '슬리피지(%)': round(slippage, 2),
            '매도일': '',
            '매도시간': '',
            '매도가격': '',
            '보유기간(일)': '',
            '수익률(%)': ''
        }

        try:
            df_new = pd.DataFrame([new_data])
            if not os.path.exists(self.CSV_FILE_NAME):
                df_new.to_csv(self.CSV_FILE_NAME, index=False, encoding='utf-8-sig')
            else:
                df_new.to_csv(self.CSV_FILE_NAME, mode='a', header=False, index=False, encoding='utf-8-sig')
            print(f"[기록] 매수 이력 저장 완료: {stock_name}")
        except Exception as e:
            print(f"[오류] 매수 기록 저장 실패: {e}")

    def _log_sell_trade(self, code, stock_name, sell_date, sell_time, sell_price):
        """매도 체결 시: 기존 행을 찾아 매도 정보 업데이트"""
        if not os.path.exists(self.CSV_FILE_NAME):
            print(f"[오류] 매매 기록 파일이 없어 매도 업데이트 실패: {stock_name}")
            return

        try:
            # 종목코드를 문자열로 읽어야 '005930' 등이 유지됨
            df = pd.read_csv(self.CSV_FILE_NAME, dtype={'종목코드': str})

            # 해당 종목이면서 '매도일'이 비어있는 행 찾기 (현재 보유 중인 건)
            mask = (df['종목코드'] == code) & (df['매도일'].isna() | (df['매도일'] == ''))

            if not df.loc[mask].empty:
                # 가장 최근 매수 건(마지막 인덱스) 선택
                idx = df.loc[mask].index[-1]

                # 수익률 및 보유기간 계산을 위해 기존 데이터 로드
                buy_price = float(df.loc[idx, '실제매입가'])
                buy_date = str(df.loc[idx, '매수일'])

                # 보유 기간 (Business Day 기준)
                try:
                    hold_days = np.busday_count(buy_date, sell_date)
                except:
                    hold_days = 0

                    # 수익률
                return_rate = ((sell_price - buy_price) / buy_price) * 100

                # 데이터 업데이트
                df.loc[idx, '매도일'] = sell_date
                df.loc[idx, '매도시간'] = sell_time
                df.loc[idx, '매도가격'] = sell_price
                df.loc[idx, '보유기간(일)'] = hold_days
                df.loc[idx, '수익률(%)'] = round(return_rate, 2)

                # 파일 덮어쓰기
                df.to_csv(self.CSV_FILE_NAME, index=False, encoding='utf-8-sig')
                print(f"[기록] 매도 이력 업데이트 완료: {stock_name} (수익률: {return_rate:.2f}%)")
            else:
                print(f"[알림] '{stock_name}'의 매수 기록을 찾을 수 없어 매도 기록만 별도로 남길 수 없습니다.")

        except Exception as e:
            print(f"[오류] 매도 기록 업데이트 실패: {e}")

    # -------------------------------------
    # 초기화 체인
    # -------------------------------------
    def _req_outstanding_orders(self):
        print("[시스템] 미체결 주문 내역을 확인합니다...")
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "전체종목구분", "0")
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "매매구분", "0")
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "체결구분", "1")
        self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "미체결요청", "opt10075", 0, "0102")

    def _req_account_balance(self):
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

                # 종목명, 수량, 매입가 확인
                name = self.kiwoom.dynamicCall("GetMasterCodeName(QString)", code)
                qty = int(self.kiwoom.dynamicCall("GetCommData(QString, QString, int, QString)", trcode, rqname, i,
                                                  "보유수량").strip())
                price = int(self.kiwoom.dynamicCall("GetCommData(QString, QString, int, QString)", trcode, rqname, i,
                                                    "매입가").strip())

                if qty > 0:
                    # 목표가 정확히 계산 (전일종가 * 1.05)
                    prev_price_str = self.kiwoom.dynamicCall("GetMasterLastPrice(QString)", code)
                    prev_price = abs(int(prev_price_str)) if prev_price_str else 0

                    target_raw = int(prev_price * 1.05) if prev_price > 0 else price

                    # 매수일시 고정
                    buy_date = '2026-02-19'
                    buy_time = '09:00:00'

                    self.held_stocks[code] = {
                        'qty': qty, 'price': price, 'buy_date': buy_date, 'buy_time': buy_time,
                        'target_raw': target_raw, 'type': '지정가'
                    }

                    # [수정] 잔고 조회 시에는 CSV에 저장하지 않음 (기존 보유분 제외)
                    # self._log_buy_trade(...)  <- 삭제됨

            print(f"[시스템] 보유 종목 리스트 복원: {len(self.held_stocks)}종목")
            self._print_slippage_report()
            QTimer.singleShot(200, self._get_condition_load)

    # -------------------------------------
    # 메인 로직
    # -------------------------------------
    def _execute_buy(self, code):
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

        prev_price = abs(int(self.kiwoom.dynamicCall("GetMasterLastPrice(QString)", code)))
        if prev_price == 0: return

        raw_target_price = prev_price * 1.05
        target_price = self._adjust_price_to_tick(raw_target_price)
        quantity = 1000000 // target_price

        if quantity == 0:
            print(f"[{now}] [매수불가] {stock_name}({code}) - 단가 초과")
            return

        print(f"[{now}] [자동매수] {stock_name}({code}) | 목표가: {target_price:,}원")

        # 메타 데이터 저장
        self.buy_meta_data[code] = {'target_raw': raw_target_price, 'time': now}
        self._send_order(code, 1, quantity, target_price)

        if code not in self.bought_today: self.bought_today.append(code)

    def _execute_sell(self, code):
        now = datetime.now().strftime('%H:%M:%S')
        stock_name = self.kiwoom.dynamicCall("GetMasterCodeName(QString)", code)

        if code in self.bought_today:
            print(f"[{now}] [매도스킵] {stock_name} - 당일 매수 종목")
            return

        if code in self.held_stocks and self.held_stocks[code]['qty'] > 0:
            quantity = self.held_stocks[code]['qty']
            print(f"[{now}] [자동매도] {stock_name} {quantity}주 시장가 매도")
            self._send_order(code, 2, quantity, 0)
        else:
            print(f"[{now}] [매도불가] {stock_name} - 잔고 없음")

    # -------------------------------------
    # 이벤트 핸들러 (연결/주문/체결)
    # -------------------------------------
    def comm_connect(self):
        print("[시스템] 로그인 시도...")
        self.kiwoom.dynamicCall("CommConnect()")
        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_()

    def _event_connect(self, err_code):
        if err_code == 0:
            print("로그인 성공!")
            self._get_account_info()
            QTimer.singleShot(200, self._req_outstanding_orders)
        else:
            print("로그인 실패")
            self.login_event_loop.exit()

    def _handler_msg(self, scr_no, rqname, trcode, msg):
        if "매수" in rqname or "주문" in msg: print(f"[서버메시지] {msg}")

    def after_login(self):
        self.login_event_loop.exit()

    def _get_account_info(self):
        self.account_num = self.kiwoom.dynamicCall("GetLoginInfo(QString)", "ACCNO").split(';')[0]
        print(f"[내 정보] 계좌번호: {self.account_num}")

    def _get_condition_load(self):
        self.kiwoom.dynamicCall("GetConditionLoad()")

    def _handler_condition_load(self, ret, msg):
        if ret == 1:
            print("[시스템] 조건식 로딩 완료.")
            conditions = self.kiwoom.dynamicCall("GetConditionNameList()").split(";")[:-1]
            for c in conditions:
                idx, name = c.split('^')
                if name in [self.BUY_STRATEGY_NAME, self.SELL_STRATEGY_NAME]:
                    print(f"[{name}] 초기 검색 요청...")
                    self.kiwoom.dynamicCall("SendCondition(QString, QString, int, int)",
                                            "0156" if "매수" in name else "0157", name, int(idx), 0)

    def _handler_condition(self, scr_no, code_list, cond_name, cond_index, next):
        now = datetime.now().strftime('%H:%M:%S')
        codes = code_list.split(';')[:-1] if code_list else []
        print(f"\n[{now}] [{cond_name}] 검색 결과: {len(codes)}종목")

        for code in codes:
            if cond_name.strip() == self.BUY_STRATEGY_NAME.strip():
                self.current_conditioned_stocks.add(code)
                self._execute_buy(code)
            elif cond_name.strip() == self.SELL_STRATEGY_NAME.strip():
                self._execute_sell(code)
            time.sleep(0.3)

        print(f"[시스템] {cond_name} 실시간 감시 전환")
        self.kiwoom.dynamicCall("SendCondition(QString, QString, int, int)", scr_no, cond_name, int(cond_index), 1)

    def _handler_real_condition(self, code, type, cond_name, cond_index):
        if type == 'I':
            if cond_name.strip() == self.BUY_STRATEGY_NAME.strip():
                self.current_conditioned_stocks.add(code)
                self._execute_buy(code)
            elif cond_name.strip() == self.SELL_STRATEGY_NAME.strip():
                self._execute_sell(code)
        elif type == 'D' and cond_name.strip() == self.BUY_STRATEGY_NAME.strip():
            self.current_conditioned_stocks.discard(code)

    def _send_order(self, code, order_type, quantity, price):
        hoga = "03" if order_type == 2 else "00"
        self.kiwoom.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                                ["send_order", "0101", self.account_num, order_type, code, quantity, price, hoga, ""])

    def _handler_chejan_data(self, gubun, item_cnt, fid_list):
        if gubun == '0':  # 접수/체결
            status = self.kiwoom.dynamicCall("GetChejanData(int)", 913)
            code = self.kiwoom.dynamicCall("GetChejanData(int)", 9001).replace('A', '')
            order_no = self.kiwoom.dynamicCall("GetChejanData(int)", 9203)
            order_type = self.kiwoom.dynamicCall("GetChejanData(int)", 905)

            if "매수" in order_type:
                if status == "접수":
                    self.open_buy_orders[code] = order_no
                elif status == "체결":
                    if code in self.open_buy_orders: del self.open_buy_orders[code]

                    if code not in self.held_stocks:
                        buy_price = int(self.kiwoom.dynamicCall("GetChejanData(int)", 910))
                        stock_name = self.kiwoom.dynamicCall("GetMasterCodeName(QString)", code)
                        today = datetime.now().strftime('%Y-%m-%d')
                        now_time = datetime.now().strftime('%H:%M:%S')

                        target_raw = 0
                        if code in self.buy_meta_data: target_raw = self.buy_meta_data[code]['target_raw']

                        slippage = 0
                        if target_raw > 0:
                            slippage = ((buy_price - target_raw) / target_raw) * 100

                        self._log_buy_trade(code, stock_name, today, now_time, target_raw, buy_price, slippage)

                        self.held_stocks[code] = {
                            'qty': 0, 'price': buy_price, 'buy_date': today, 'buy_time': now_time,
                            'target_raw': target_raw, 'type': '지정가'
                        }

            elif "매도" in order_type and status == "체결":
                sell_price = int(self.kiwoom.dynamicCall("GetChejanData(int)", 910))
                stock_name = self.kiwoom.dynamicCall("GetMasterCodeName(QString)", code)
                today = datetime.now().strftime('%Y-%m-%d')
                now_time = datetime.now().strftime('%H:%M:%S')

                self._log_sell_trade(code, stock_name, today, now_time, sell_price)

            print(f"[체결알림] {code} | {status} | {order_no}")

        elif gubun == '1':  # 잔고통보
            code = self.kiwoom.dynamicCall("GetChejanData(int)", 9001).replace('A', '')
            qty = int(self.kiwoom.dynamicCall("GetChejanData(int)", 930))
            if qty > 0:
                if code in self.held_stocks:
                    self.held_stocks[code]['qty'] = qty
                else:
                    self.held_stocks[code] = {'qty': qty, 'price': 0}
            else:
                if code in self.held_stocks: del self.held_stocks[code]

    def _periodic_check(self):
        now = datetime.now()
        print(f"\n[시스템 점검] {now.strftime('%H:%M:%S')} 실시간 감시 작동 중...")

        # [수정] 실시간 검색된 종목 수 명확하게 표시 (len 사용)
        search_count = len(self.current_conditioned_stocks)
        print(f"  > [실시간 검색] 현재 매수 조건 포착 종목 수: {search_count}개")

        overlap_held = len(self.current_conditioned_stocks.intersection(set(self.held_stocks.keys())))
        overlap_order = len(self.current_conditioned_stocks.intersection(set(self.open_buy_orders.keys())))

        print(f"  > [계좌 현황] 보유종목: {len(self.held_stocks)}개, 미체결주문: {len(self.open_buy_orders)}건")

        if overlap_held > 0 or overlap_order > 0:
            print(f"  > (참고) 검색 종목 중 {overlap_held}개는 이미 보유 중, {overlap_order}개는 매수 진행 중")

        if now.hour == 15 and now.minute >= 20 and self.open_buy_orders:
            print("[장마감] 미체결 취소")
            for code, order_no in list(self.open_buy_orders.items()):
                self._send_order(code, 3, 0, 0)
                del self.open_buy_orders[code]
                time.sleep(0.2)

    def _print_slippage_report(self):
        if not self.held_stocks: return
        print(f"\n[시스템] 슬리피지 분석 ({datetime.now().strftime('%H:%M:%S')})")
        print("-" * 100)
        print(f"{'종목명':<10} | {'매수일시':<20} | {'목표가':<10} | {'매입가':<10} | {'슬리피지%':<10}")
        print("-" * 100)

        total_slippage = 0
        count = 0

        for code, info in self.held_stocks.items():
            name = self.kiwoom.dynamicCall("GetMasterCodeName(QString)", code)
            target = info.get('target_raw', 0)
            slippage = ((info['price'] - target) / target * 100) if target > 0 else 0

            total_slippage += slippage
            count += 1

            print(
                f"{name:<10} | {info.get('buy_date', '')} {info.get('buy_time', '')} | {int(target):<10} | {info['price']:<10} | {slippage:.2f}%")

        print("-" * 100)
        if count > 0:
            avg_slippage = total_slippage / count
            print(f"[전체 평균 슬리피지] {avg_slippage:.2f}%")
        print("-" * 100)
