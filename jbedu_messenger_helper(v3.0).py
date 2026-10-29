# -*- coding: utf-8 -*-
"""
전북특별자치도교육청 메신저 작업 도우미 v3.0
- 검색: 필터 '전체' 고정, 검색어1 → 검색어2(결과 내 재검색) → 여러 명이면 검색어3
        찾은 사람 정보에 쓴 검색어가 모두 있는지 최종 확인 (이전 결과·다른 사람 방지)
- 사용자그룹 추가: 작업용 '!!!사용자그룹추가'에 넣고 끝나면 이름 바꾸기, 이미 있으면 보관/이어서
- 쪽지 발송: 엑셀/CSV 명단, {열이름} 치환, 서식 쪽지, 서식별 첨부, 시험 발송, 미리보기,
            검색만 확인, 발송 기록(이어 보내기·중복 방지), 본문 입력·전송 완료 확인
- 쪽지 서식: 기본 쪽지 하나로 시작, 서식은 자유롭게 추가·수정(쪽지서식.json에 저장)
- 공통: F12는 어느 창이 앞에 있어도 중지, 메신저가 앞에서 사라지면 자동으로 멈춤, logs 폴더에 작업 기록, 진단 기능
"""
"""
JBEdu Messenger+ 사용자그룹 첫 번째 그룹 생성 + 추가 실제 완료 테스트 v48 FAST TWO-CHECK

목적:
- 실행 시작 시 사용자그룹 탭에서 '!!!사용자그룹추가' 그룹을 먼저 생성한다.
- 이미 같은 이름이 있으면 '동일한 이름이 존재합니다.' 메시지를 감지하고 즉시 종료한다.
- 대상자 검증은 이름+부서 재검색 후 결과 리스트 키보드 순회 방식으로 수행한다.
- 검증 성공자만 VK_APPS로 컨텍스트 메뉴를 열고 '내목록 추가'를 실행한다.
- 사용자그룹 선택창에서 다른 키보드 시도 없이 조직도 트리 첫 번째 행을 바로 선택한다.

원칙:
- 메신저 메인창 위치 변경 안 함
- 사용자그룹 선택창 위치 변경 안 함
- 메신저 메인창/검색 결과에서는 마우스 좌표 클릭 안 함
- 사용자그룹 선택창의 첫 번째 그룹은 조직도 트리 첫 행 메시지 선택으로 바로 처리함
- OCR 안 함
- WM_COMMAND 직접 실행 안 함
- WM_CONTEXTMENU 안 함
- Shift+F10 안 함
- 컨텍스트 메뉴는 VK_APPS 단일 방식만 사용
- 메뉴가 실제로 떠 있을 때만 ESC 사용

전제:
- '!!!사용자그룹추가' 그룹이 없으면 이 스크립트가 생성을 시도한다.
- 이미 있으면 중복 메시지를 감지하고 전체 작업을 종료한다.

주의:
- ADD 입력 시 실제 그룹 생성 및 실제 사용자그룹 추가가 수행될 수 있음
"""

from pywinauto import Desktop, Application, keyboard
from datetime import datetime
from contextlib import redirect_stdout
import win32gui
import win32con
import win32clipboard
import time
import os
import re
import traceback


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
AUTO_GROUP_NAME = "!!!사용자그룹추가"
STATE_SYSTEM_FOCUSED = 0x00000004


class CFG:
    """
    내목록 추가 팝업에서 첫 번째 그룹을 선택하기 위한 설정.

    v46 기준:
    - v37의 사용자그룹 탭 자동 초기화/광범위 좌클릭 fallback은 제거한다.
    - 사용자그룹 생성 단계에서 마우스 좌표 클릭을 사용하지 않는다.
    - 사용자그룹 탭 내부 컨트롤에 포커스를 준 뒤 LEFT 기준으로 그룹 메뉴를 연다.
    - 사용자 개인 메뉴가 뜨면 닫고, HOME 없이 LEFT 계열만 재시도한다.

    기존 v46 변경점:
    - 내목록 추가 선택창에서는 키보드 방식 재시도를 하지 않는다.
    - 조직도 트리 첫 번째 행을 Win32 메시지로 바로 선택한다.
    - 실제 마우스 커서는 이동하지 않는다.
    """
    ENABLE_MOUSE_FALLBACK_FOR_FIRST_GROUP = False
    FIRST_GROUP_CLIENT_X = None
    FIRST_GROUP_CLIENT_Y = None
    FIRST_GROUP_TREE_OFFSET_X = 28
    FIRST_GROUP_TREE_OFFSET_Y = 18
    AFTER_FIRST_GROUP_CLICK_DELAY = 0.35
    AFTER_CONFIRM_WAIT = 0.35
    VERBOSE_MENU = False
    VERBOSE_RESULT_SCAN = False
    MENU_OPEN_WAIT = 0.28
    MENU_ACTION_WAIT = 0.25
    SEARCH_WAIT = 0.85
    SHORT_WAIT = 0.06
    INPUT_WAIT = 0.12



# ----------------------------------------------------------------------
# 기본 유틸
# ----------------------------------------------------------------------

def safe_text(value):
    try:
        if value is None:
            return ""
        return str(value).replace("\r", " ").replace("\n", " ").strip()
    except Exception:
        return ""


def compact_text(text):
    return " ".join(safe_text(text).split())


def normalize(text):
    return re.sub(r"\s+", "", safe_text(text)).lower()


def info_of(ctrl):
    if ctrl is None:
        return "없음"

    info = ctrl.element_info

    try:
        rect = ctrl.rectangle()
        rect_str = str(rect)
    except Exception:
        rect_str = ""

    return {
        "name": safe_text(info.name),
        "control_type": safe_text(info.control_type),
        "class_name": safe_text(info.class_name),
        "auto_id": safe_text(info.automation_id),
        "handle": getattr(info, "handle", None),
        "rect": rect_str,
    }


def hwnd_info(hwnd):
    if not hwnd:
        return "없음"

    try:
        return {
            "hwnd": hwnd,
            "title": safe_text(win32gui.GetWindowText(hwnd)),
            "class_name": safe_text(win32gui.GetClassName(hwnd)),
            "rect": win32gui.GetWindowRect(hwnd),
            "visible": bool(win32gui.IsWindowVisible(hwnd)),
            "enabled": bool(win32gui.IsWindowEnabled(hwnd)),
        }
    except Exception as e:
        return f"hwnd_info 실패: {e}"


def wm_gettext(hwnd):
    try:
        length = win32gui.SendMessage(hwnd, win32con.WM_GETTEXTLENGTH, 0, 0)
        buffer = win32gui.PyMakeBuffer((length + 1) * 2)
        win32gui.SendMessage(hwnd, win32con.WM_GETTEXT, length + 1, buffer)
        raw = buffer[:]
        try:
            return raw.decode("utf-16le").rstrip("\x00")
        except Exception:
            return safe_text(win32gui.GetWindowText(hwnd))
    except Exception:
        try:
            return safe_text(win32gui.GetWindowText(hwnd))
        except Exception:
            return ""


def is_visible_rect(ctrl):
    try:
        r = ctrl.rectangle()
        if r.left < -10000 or r.top < -10000:
            return False
        if r.right <= r.left or r.bottom <= r.top:
            return False
        return True
    except Exception:
        return False


def rect_width(ctrl):
    try:
        r = ctrl.rectangle()
        return int(r.right - r.left)
    except Exception:
        return 0


def rect_top(ctrl):
    try:
        return int(ctrl.rectangle().top)
    except Exception:
        return 0


def set_clipboard_text(text):
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, text)
    finally:
        win32clipboard.CloseClipboard()



def focus_and_replace_text(edit, text, label):
    """빠른 입력: 포커스 → 전체선택 → 붙여넣기."""
    print(f"\n[{label} 입력] {text}")
    edit.set_focus()
    time.sleep(CFG.SHORT_WAIT)
    keyboard.send_keys("^a")
    time.sleep(0.03)
    keyboard.send_keys("{BACKSPACE}")
    time.sleep(0.03)
    set_clipboard_text(text)
    keyboard.send_keys("^v")
    time.sleep(CFG.INPUT_WAIT)
    print(f"{label} 입력 완료")


def invoke_control(ctrl, label):
    print(f"\n[실행] {label}")
    try:
        ctrl.invoke()
        time.sleep(CFG.MENU_ACTION_WAIT)
        print("invoke 성공")
        return True
    except Exception as e:
        print(f"invoke 실패: {e}")

    try:
        ctrl.set_focus()
        time.sleep(CFG.SHORT_WAIT)
        keyboard.send_keys("{ENTER}")
        time.sleep(CFG.MENU_ACTION_WAIT)
        print("set_focus + Enter 성공")
        return True
    except Exception as e:
        print(f"set_focus + Enter 실패: {e}")

    try:
        ctrl.set_focus()
        time.sleep(CFG.SHORT_WAIT)
        keyboard.send_keys("{SPACE}")
        time.sleep(CFG.MENU_ACTION_WAIT)
        print("set_focus + Space 성공")
        return True
    except Exception as e:
        print(f"set_focus + Space 실패: {e}")

    return False

# ----------------------------------------------------------------------
# 메신저 창 / 컨트롤 탐색
# ----------------------------------------------------------------------

def find_messenger_hwnd():
    for win in Desktop(backend="uia").windows():
        try:
            d = info_of(win)
            if d["name"] == "JBEdu Messenger+" or d["class_name"] == "@Messenger7_MainWnd":
                return d["handle"]
        except Exception:
            pass

    candidates = []

    def enum_callback(hwnd, _):
        try:
            title = safe_text(win32gui.GetWindowText(hwnd))
            class_name = safe_text(win32gui.GetClassName(hwnd))

            if title == "JBEdu Messenger+" or class_name == "@Messenger7_MainWnd":
                candidates.append(hwnd)
        except Exception:
            pass

        return True

    try:
        win32gui.EnumWindows(enum_callback, None)
    except Exception:
        pass

    return candidates[0] if candidates else None


def restore_messenger_without_move(hwnd):
    """
    창 위치/크기 변경 금지.
    최소화 상태면 복원만 하고, 가능한 경우 포커스만 준다.
    """
    try:
        win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
        time.sleep(0.07)

        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            time.sleep(0.07)

        try:
            win32gui.SetForegroundWindow(hwnd)
        except Exception as e:
            print(f"SetForegroundWindow 실패, 무시하고 계속 진행: {e}")

        time.sleep(0.18)

    except Exception as e:
        print(f"메신저 창 복원 실패: {e}")


def get_root(hwnd):
    app = Application(backend="uia").connect(handle=hwnd)
    return app.window(handle=hwnd).wrapper_object()


def find_controls(root, control_type=None, name=None, class_name=None, auto_id=None):
    result = []

    try:
        controls = root.descendants()
    except Exception:
        return result

    for ctrl in controls:
        try:
            d = info_of(ctrl)

            if not is_visible_rect(ctrl):
                continue

            if control_type is not None and d["control_type"] != control_type:
                continue

            if name is not None and d["name"] != name:
                continue

            if class_name is not None and d["class_name"] != class_name:
                continue

            if auto_id is not None and d["auto_id"] != str(auto_id):
                continue

            result.append(ctrl)

        except Exception:
            pass

    return result


# ----------------------------------------------------------------------
# 메뉴 공통: VK_APPS 단일 방식 + 키보드 포커스 선택
# ----------------------------------------------------------------------

def find_context_menu():
    for win in Desktop(backend="uia").windows():
        try:
            d = info_of(win)
            if d["control_type"] == "Menu" and d["class_name"] == "#32768":
                return win
        except Exception:
            pass

    return None



def close_context_menu_only():
    menu = find_context_menu()
    if menu is None:
        return

    try:
        keyboard.send_keys("{ESC}")
        time.sleep(CFG.SHORT_WAIT)
        print("컨텍스트 메뉴 닫음")
    except Exception as e:
        print(f"컨텍스트 메뉴 닫기 실패: {e}")


def get_menu_items(menu):
    rows = []

    try:
        items = menu.descendants()
    except Exception as e:
        print(f"메뉴 항목 조회 실패: {e}")
        return rows

    if CFG.VERBOSE_MENU:
        print("\n" + "=" * 80)
        print("[컨텍스트 메뉴 항목]")
        print("=" * 80)
        print(info_of(menu))

    for idx, item in enumerate(items):
        try:
            d = info_of(item)
            if d["control_type"] == "MenuItem":
                legacy = {}
                try:
                    legacy = item.legacy_properties()
                except Exception:
                    legacy = {}

                state = 0
                try:
                    state = int(legacy.get("State", 0))
                except Exception:
                    state = 0

                if CFG.VERBOSE_MENU:
                    print(
                        f"{idx:03d} | name={d['name']} | auto_id={d['auto_id']} | "
                        f"state={state} | rect={d['rect']}"
                    )

                rows.append({"ctrl": item, "info": d, "state": state})
        except Exception:
            pass

    return rows

def get_focused_menu_item_name(menu):
    rows = get_menu_items(menu)

    for row in rows:
        name = row["info"]["name"]
        state = row["state"]

        has_focus = False
        try:
            has_focus = bool(row["ctrl"].has_keyboard_focus())
        except Exception:
            pass

        if has_focus or (state & STATE_SYSTEM_FOCUSED):
            return name

    return ""


def menu_item_name_matches(focused_name, target_names):
    focused_name = safe_text(focused_name)

    if isinstance(target_names, str):
        target_names = [target_names]

    focused_norm = normalize(focused_name)

    for target in target_names:
        target = safe_text(target)
        target_norm = normalize(target)

        if not target_norm:
            continue

        # 1순위: 정확히 일치
        if focused_norm == target_norm:
            return True

        # 2순위: 메뉴명에 목표 문구가 포함되는 경우
        # 예: '사용자그룹 추가...'처럼 뒤에 말줄임표/단축키가 붙는 경우 대응
        if len(target_norm) >= 5 and target_norm in focused_norm:
            return True

    return False



def select_menu_item_by_keyboard(menu, target_name):
    """
    빠른 메뉴 선택.
    1순위: UIA MenuItem.invoke()
    2순위: 한 번만 메뉴 목록을 읽고 목표 위치까지 DOWN 후 ENTER
    기존처럼 매 step마다 메뉴 전체를 재출력/재스캔하지 않는다.
    """
    targets = target_name if isinstance(target_name, (list, tuple, set)) else [target_name]
    targets = [safe_text(t) for t in targets]
    print(f"\n[메뉴 선택] {', '.join(targets)}")

    rows = get_menu_items(menu)
    visible_rows = [r for r in rows if safe_text(r["info"].get("name"))]
    names = [r["info"].get("name", "") for r in visible_rows]
    print(f"메뉴명: {names}")

    target_idx = None
    target_row = None
    for idx, row in enumerate(visible_rows):
        name = row["info"].get("name", "")
        if menu_item_name_matches(name, targets):
            target_idx = idx
            target_row = row
            break

    if target_row is None:
        print(f"목표 메뉴 항목을 찾지 못했습니다: {targets}")
        return False

    target_label = target_row["info"].get("name", "")

    # UIA invoke가 되면 가장 빠르다.
    try:
        target_row["ctrl"].invoke()
        time.sleep(CFG.MENU_ACTION_WAIT)
        print(f"{target_label} invoke 실행 완료")
        return True
    except Exception as e:
        print(f"{target_label} invoke 실패, 키보드 방식으로 전환: {e}")

    try:
        keyboard.send_keys("{HOME}")
        time.sleep(0.06)
        # 이 메신저 메뉴는 HOME 직후 첫 DOWN에서 첫 번째 유효 항목으로 들어가는 경우가 많다.
        # target_idx가 0이면 DOWN 1회, 1이면 DOWN 2회가 필요했다.
        downs = target_idx + 1
        if downs > 0:
            keyboard.send_keys("{DOWN}" * downs)
            time.sleep(0.04 * downs)
        keyboard.send_keys("{ENTER}")
        time.sleep(CFG.MENU_ACTION_WAIT)
        print(f"{target_label} 키보드 실행 완료")
        return True
    except Exception as e:
        print(f"{target_label} 키보드 실행 실패: {e}")
        return False


def open_context_menu_on_control(ctrl, label):
    print(f"\n[컨텍스트 메뉴 열기] {label}")
    close_context_menu_only()

    try:
        ctrl.set_focus()
        time.sleep(0.10)
    except Exception as e:
        print(f"대상 컨트롤 set_focus 실패: {e}")
        return None

    try:
        keyboard.send_keys("{VK_APPS}")
        time.sleep(CFG.MENU_OPEN_WAIT)
    except Exception as e:
        print(f"VK_APPS 전송 실패: {e}")
        return None

    menu = find_context_menu()
    if menu is None:
        print("컨텍스트 메뉴를 찾지 못했습니다.")
        return None

    print("컨텍스트 메뉴 열림")
    return menu

# ----------------------------------------------------------------------
# 초기 사용자그룹 생성
# ----------------------------------------------------------------------

def click_usergroup_tab(root):
    buttons = find_controls(
        root,
        control_type="Button",
        name="사용자그룹",
        auto_id="1011",
    )

    if not buttons:
        print("사용자그룹 탭을 찾지 못했습니다.")
        return False

    return invoke_control(buttons[0], "상단 사용자그룹 탭")


def find_usergroup_main_list(root):
    # 사용자그룹 탭의 주 목록은 기존 진단상 '내목록' UltariListScrollView
    lists = find_controls(
        root,
        control_type="Pane",
        name="내목록",
        class_name="UltariListScrollView",
    )

    if lists:
        return lists[0]

    # fallback: 보이는 UltariListScrollView 중 '내목록' 텍스트가 있는 것
    try:
        controls = root.descendants()
    except Exception:
        controls = []

    candidates = []
    for ctrl in controls:
        try:
            d = info_of(ctrl)
            if not is_visible_rect(ctrl):
                continue
            if d["class_name"] != "UltariListScrollView":
                continue
            text = " ".join([d["name"], safe_text(ctrl.window_text())])
            if "내목록" in text:
                candidates.append(ctrl)
        except Exception:
            pass

    return candidates[0] if candidates else None


def collect_top_window_handles():
    handles = set()

    def enum_callback(hwnd, _):
        try:
            if win32gui.IsWindow(hwnd):
                handles.add(hwnd)
        except Exception:
            pass
        return True

    try:
        win32gui.EnumWindows(enum_callback, None)
    except Exception:
        pass

    try:
        for win in Desktop(backend="uia").windows():
            try:
                h = info_of(win).get("handle")
                if h:
                    handles.add(h)
            except Exception:
                pass
    except Exception:
        pass

    return handles


def collect_texts_from_window_handle(hwnd):
    texts = []

    def add_text(value):
        value = safe_text(value)
        if value and value not in texts:
            texts.append(value)

    try:
        add_text(win32gui.GetWindowText(hwnd))
    except Exception:
        pass

    try:
        add_text(wm_gettext(hwnd))
    except Exception:
        pass

    try:
        ctrl = Desktop(backend="uia").window(handle=hwnd).wrapper_object()
        d = info_of(ctrl)
        add_text(d.get("name"))

        try:
            add_text(ctrl.window_text())
        except Exception:
            pass

        try:
            for t in ctrl.texts():
                add_text(t)
        except Exception:
            pass

        try:
            legacy = ctrl.legacy_properties()
            if legacy:
                for _, v in legacy.items():
                    add_text(v)
        except Exception:
            pass

        try:
            for child in ctrl.descendants():
                cd = info_of(child)
                add_text(cd.get("name"))
                try:
                    add_text(child.window_text())
                except Exception:
                    pass
                try:
                    for t in child.texts():
                        add_text(t)
                except Exception:
                    pass
                try:
                    legacy = child.legacy_properties()
                    if legacy:
                        for _, v in legacy.items():
                            add_text(v)
                except Exception:
                    pass
        except Exception:
            pass

    except Exception:
        pass

    return texts


def collect_control_names_from_window_handle(hwnd):
    names = []

    def add(value):
        value = safe_text(value)
        if value and value not in names:
            names.append(value)

    try:
        ctrl = Desktop(backend="uia").window(handle=hwnd).wrapper_object()
        add(info_of(ctrl).get("name"))

        try:
            for child in ctrl.descendants():
                d = info_of(child)
                add(d.get("name"))
        except Exception:
            pass
    except Exception:
        pass

    return names


def collect_button_names_from_window_handle(hwnd):
    names = []

    try:
        ctrl = Desktop(backend="uia").window(handle=hwnd).wrapper_object()
        for child in ctrl.descendants():
            d = info_of(child)
            if d["control_type"] == "Button" and d["name"]:
                if d["name"] not in names:
                    names.append(d["name"])
    except Exception:
        pass

    return names


def count_edit_controls_from_window_handle(hwnd):
    count = 0

    try:
        ctrl = Desktop(backend="uia").window(handle=hwnd).wrapper_object()
        for child in ctrl.descendants():
            d = info_of(child)
            if d["control_type"] == "Edit":
                count += 1
    except Exception:
        pass

    return count


def looks_like_duplicate_group_popup(hwnd, combined_text):
    """
    그룹명 입력 직후 뜬 JBEdu Messenger+ 팝업을 판단한다.

    v34 핵심:
    - 메신저가 '동일한 이름이 존재합니다.' 문구와 '확인' 버튼명을 UIA로 숨기는 경우가 있다.
    - 그래서 문구가 읽히지 않더라도, 그룹명 입력 직후 작은 @Messenger7_Wnd 팝업이 남아 있으면
      성공으로 보지 않고 즉시 종료 대상으로 처리한다.
    - 즉, 여기서는 '중복'뿐 아니라 '그룹 생성을 막는 메시지창'도 모두 DUPLICATE/BLOCKED로 본다.
    """
    needles = [
        "동일한 이름이 존재합니다",
        "동일한 이름",
        "이미 존재",
    ]

    for n in needles:
        if n in combined_text:
            return True, f"문구 감지: {n}"

    try:
        title = safe_text(win32gui.GetWindowText(hwnd))
        class_name = safe_text(win32gui.GetClassName(hwnd))
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        width = int(right - left)
        height = int(bottom - top)
    except Exception:
        title = ""
        class_name = ""
        width = 0
        height = 0

    button_names = collect_button_names_from_window_handle(hwnd)
    edit_count = count_edit_controls_from_window_handle(hwnd)
    control_names = collect_control_names_from_window_handle(hwnd)

    if title == "JBEdu Messenger+ Frame Window" and class_name == "@Messenger7_Wnd":
        # 1) 정상적으로 확인 버튼명이 읽히는 메시지창
        if "확인" in button_names and "취소" not in button_names and edit_count == 0:
            return True, "JBEdu Messenger+ 확인 단독 메시지창 감지"

        # 2) 이번 로그처럼 버튼명/문구가 전혀 읽히지 않는 작은 메시지창
        #    예: rect 폭 256, 높이 117 / button_names=[] / edit_count=0
        if edit_count == 0 and width <= 520 and height <= 260:
            return True, f"JBEdu Messenger+ 텍스트 비노출 메시지창 감지(width={width}, height={height})"

        # 3) 텍스트는 안 읽히지만 컨트롤 이름에 문구가 잡히는 경우
        combined_names = " ".join(control_names + button_names)
        if "동일한 이름" in combined_names or "이미 존재" in combined_names:
            return True, "컨트롤 이름에서 중복 문구 감지"

    combined_names = " ".join(control_names + button_names)
    if "동일한 이름" in combined_names or "이미 존재" in combined_names:
        return True, "컨트롤 이름에서 중복 문구 감지"

    return False, ""


def wait_duplicate_group_message(before_handles, main_hwnd=None, timeout=3.5):
    """
    그룹명 입력 후 중복 메시지를 감지한다.

    v34 수정점:
    - 새로 뜬 창만 보지 않고 현재 foreground 창도 항상 검사한다.
    - JBEdu Messenger+ Frame Window가 텍스트를 비워 둔 채 '확인' 단독 메시지창으로 뜨는 경우도
      '동일한 이름이 존재합니다.' 상황으로 간주한다.
    - 감지만 하고 닫지는 않는다. 실제 종료는 main 흐름에서 처리한다.
    """
    end = time.time() + timeout
    inspected_log = set()

    while time.time() < end:
        current_handles = collect_top_window_handles()
        candidate_handles = []

        for h in current_handles:
            if h not in before_handles:
                candidate_handles.append(h)

        try:
            fg = win32gui.GetForegroundWindow()
            if fg and fg not in candidate_handles:
                candidate_handles.insert(0, fg)
        except Exception:
            pass

        # JBEdu Messenger+ Frame Window가 기존 핸들을 재사용하는 경우 대비
        for h in current_handles:
            try:
                title = safe_text(win32gui.GetWindowText(h))
                class_name = safe_text(win32gui.GetClassName(h))
                if title == "JBEdu Messenger+ Frame Window" and class_name == "@Messenger7_Wnd":
                    if h not in candidate_handles:
                        candidate_handles.append(h)
            except Exception:
                pass

        for hwnd in candidate_handles:
            if not hwnd:
                continue

            if main_hwnd and hwnd == main_hwnd:
                continue

            try:
                if not win32gui.IsWindow(hwnd) or not win32gui.IsWindowVisible(hwnd):
                    continue
            except Exception:
                continue

            texts = collect_texts_from_window_handle(hwnd)
            combined = "\n".join(texts)

            is_dup, dup_reason = looks_like_duplicate_group_popup(hwnd, combined)

            if hwnd not in inspected_log:
                inspected_log.add(hwnd)
                print("\n[그룹 생성 후 창/팝업 검사]")
                print(hwnd_info(hwnd))
                print(f"button_names={collect_button_names_from_window_handle(hwnd)}")
                print(f"edit_count={count_edit_controls_from_window_handle(hwnd)}")
                if combined:
                    print(combined[:1000])
                else:
                    print("읽힌 텍스트 없음")

            if is_dup:
                print(f"중복 사용자그룹 메시지 감지: {dup_reason}")
                return hwnd, combined

        time.sleep(0.07)

    print("중복 사용자그룹 메시지 없음")
    return None, ""


def close_duplicate_message(hwnd):
    if not hwnd:
        return False

    try:
        ctrl = Desktop(backend="uia").window(handle=hwnd).wrapper_object()

        # 확인 버튼 우선
        try:
            for child in ctrl.descendants():
                d = info_of(child)
                if d["control_type"] == "Button" and d["name"] in ["확인", "OK", "예"]:
                    return invoke_control(child, "중복 메시지 확인")
        except Exception:
            pass

    except Exception:
        pass

    try:
        # 포커스가 해당 메시지창일 때만 Enter
        try:
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.07)
        except Exception:
            pass
        keyboard.send_keys("{ENTER}")
        time.sleep(0.14)
        print("중복 메시지 Enter 닫기 완료")
        return True
    except Exception as e:
        print(f"중복 메시지 닫기 실패: {e}")
        return False


def type_group_name_and_confirm(group_name, main_hwnd=None):
    print("\n" + "=" * 80)
    print("[사용자그룹명 입력]")
    print("=" * 80)
    print(f"입력할 그룹명: {group_name}")

    # 그룹명 입력으로 인해 새로 뜨는 메시지만 감지하기 위해 입력 직전 핸들 목록을 기록
    before_handles = collect_top_window_handles()

    try:
        set_clipboard_text(group_name)
        keyboard.send_keys("^v")
        time.sleep(0.10)
        keyboard.send_keys("{ENTER}")
        time.sleep(CFG.MENU_OPEN_WAIT)
        print("그룹명 입력 후 Enter 완료")
    except Exception as e:
        print(f"그룹명 입력 실패: {e}")
        return "ERROR"

    duplicate_hwnd, duplicate_text = wait_duplicate_group_message(before_handles, main_hwnd=main_hwnd, timeout=3.5)
    if duplicate_hwnd:
        print("중복 메시지는 확인을 누르지 않고 그대로 둔 채 즉시 종료 처리합니다.")
        return "DUPLICATE"

    # 중복 메시지가 새로 뜨지 않았으면 성공으로 본다.
    # 여기서 Enter를 추가로 보내지 않는다. 추가 Enter가 엉뚱한 창/메뉴에 들어갈 수 있기 때문.
    return "OK"



def menu_names_from_rows(rows):
    names = []
    for row in rows:
        try:
            name = safe_text(row["info"].get("name"))
        except Exception:
            name = ""
        if name and name not in names:
            names.append(name)
    return names


def menu_has_name(names, targets):
    if isinstance(targets, str):
        targets = [targets]
    joined = " ".join(names)
    joined_norm = normalize(joined)
    for target in targets:
        if normalize(target) in joined_norm:
            return True
    return False


def is_user_person_context_menu(names):
    """
    사용자그룹 생성 메뉴가 아니라 특정 사용자 항목에서 열린 개인 메뉴인지 판별한다.
    노트북에서 사용자그룹 탭 내부 선택이 사람 행에 남아 있으면 이런 메뉴가 열린다.
    이 경우 자동 좌클릭 초기화를 시도하지 않고 바로 중단한다.
    """
    has_person_actions = menu_has_name(names, ["대화하기", "쪽지전송", "내목록 추가"])
    has_group_actions = menu_has_name(names, ["사용자그룹 생성", "사용자그룹 추가", "하위 그룹 생성", "선택 그룹 삭제", "그룹 이름 변경"])
    return has_person_actions and not has_group_actions


def is_group_create_context_menu(names):
    return menu_has_name(names, ["사용자그룹 생성", "사용자그룹 추가"])



def try_open_group_create_menu_after_no_mouse(group_list, label):
    """
    사용자그룹 탭 안에서 '내 프로필' 항목을 기준점으로 선택한 뒤 VK_APPS 메뉴를 연다.

    데스크톱/노트북 차이 원인:
    - 같은 '내목록' 컨트롤이라도 마지막 선택이 사람 행에 남아 있으면 사용자 개인 메뉴가 열린다.
    - 사용자가 확인한 안정 기준은 사용자그룹 탭 상단의 '내 프로필' 항목 선택이다.

    반환:
    - 그룹 생성 메뉴가 확인되면 menu 반환
    - 사용자 개인 메뉴 또는 기타 메뉴면 닫고 None 반환
    """
    menu = open_context_menu_on_control(group_list, label)
    if menu is None:
        return None, "컨텍스트 메뉴 열기 실패"

    rows = get_menu_items(menu)
    names = menu_names_from_rows(rows)
    print(f"메뉴명 목록: {names}")

    if is_group_create_context_menu(names):
        print("[사용자그룹 생성 메뉴 확인] 사용자그룹 생성/추가 항목이 있습니다.")
        return menu, "OK"

    if is_user_person_context_menu(names):
        print("[검증 실패] 사용자그룹 생성 메뉴가 아니라 사용자 개인 메뉴가 열렸습니다.")
        close_context_menu_only()
        return None, "USER_MENU"

    print("[검증 실패] 사용자그룹 생성/추가 항목이 없는 메뉴입니다.")
    close_context_menu_only()
    return None, "OTHER_MENU"


def send_navigation_keys_to_group_list(group_list, key_sequence, label):
    print(f"\n[내 프로필 선택 시도 - 키보드] {label}: {key_sequence or '(키 없음)'}")
    try:
        group_list.set_focus()
        time.sleep(0.08)
        if key_sequence:
            keyboard.send_keys(key_sequence)
            time.sleep(0.05)
        return True
    except Exception as e:
        print(f"키보드 내 프로필 선택 시도 실패: {e}")
        return False


def open_group_create_menu_via_no_mouse(group_list):
    """
    v46 핵심:
    - 마우스 좌표 클릭을 전혀 사용하지 않는다.
    - 사용자 확인 결과, 사용자그룹 탭 내부에서는 HOME이 안정적으로 먹지 않는다.
    - 대신 내목록 컨트롤에 포커스를 준 뒤 방향키 입력으로 내부 선택 상태를 깨우고,
      LEFT를 눌러 그룹/root 기준으로 보낸 다음 VK_APPS를 연다.
    - 열린 메뉴에 '사용자그룹 생성/사용자그룹 추가'가 있으면 그 메뉴를 사용한다.
    - 사용자 개인 메뉴가 뜨면 닫고 다음 키보드 시퀀스를 재시도한다.
    - 끝까지 그룹 생성 메뉴가 안 뜨면 안전하게 중단한다.
    """
    print("\n" + "=" * 80)
    print("[사용자그룹 생성 메뉴 열기 - LEFT 기준 / 무좌표]")
    print("=" * 80)
    print("사용자그룹 탭 내부 컨트롤에 포커스를 준 뒤, 방향키 입력 후 LEFT → VK_APPS를 시도합니다.")
    print("※ HOME은 사용하지 않습니다. 이 버전은 마우스 좌표 클릭을 사용하지 않습니다.")

    attempts = [
        ("LEFT만", "{LEFT}"),
        ("DOWN 후 LEFT", "{DOWN}{LEFT}"),
        ("UP 후 LEFT", "{UP}{LEFT}"),
        ("RIGHT 후 LEFT", "{RIGHT}{LEFT}"),
        ("DOWN 두 번 후 LEFT", "{DOWN}{DOWN}{LEFT}"),
        ("UP 두 번 후 LEFT", "{UP}{UP}{LEFT}"),
        ("PGUP 후 LEFT", "{PGUP}{LEFT}"),
        ("CTRL+LEFT", "^{LEFT}"),
    ]

    for idx, (label, keys) in enumerate(attempts, start=1):
        print("\n" + "-" * 80)
        print(f"[LEFT 기준 메뉴 열기 시도 {idx}] {label}: {keys}")
        print("-" * 80)

        if not send_navigation_keys_to_group_list(group_list, keys, f"{idx}. {label}"):
            print("내목록 컨트롤 포커스/키 입력 실패 - 다음 방식으로 재시도")
            continue

        menu, reason = try_open_group_create_menu_after_no_mouse(
            group_list,
            f"사용자그룹 내목록 - {label} 후 LEFT 기준"
        )
        if menu is not None:
            print(f"[LEFT 기준 성공 - 무좌표] {label}")
            return menu

        if reason == "USER_MENU":
            print("사용자 개인 메뉴가 열렸으므로 닫고 다음 LEFT 방식으로 재시도합니다.")
        else:
            print(f"그룹 생성 메뉴 확인 실패: {reason}")

    print("[실패] LEFT 기준 키보드 방식만으로 사용자그룹 생성 메뉴를 열지 못했습니다.")
    print("사용자그룹 탭 안에서 아무 항목을 한 번 선택한 뒤, 왼쪽 방향키를 누르면 그룹 메뉴가 열리는지 수동으로 확인하세요.")
    return None


def create_auto_group_before_run(main_hwnd):
    print("\n" + "=" * 80)
    print(f"[초기 사용자그룹 생성: {AUTO_GROUP_NAME}]")
    print("=" * 80)

    root = get_root(main_hwnd)

    if not click_usergroup_tab(root):
        return "ERROR", "사용자그룹 탭 클릭 실패"

    time.sleep(CFG.MENU_OPEN_WAIT)
    root = get_root(main_hwnd)

    group_list = find_usergroup_main_list(root)
    if group_list is None:
        return "ERROR", "사용자그룹 내목록 리스트 찾기 실패"

    print("사용자그룹 내목록 리스트:")
    print(info_of(group_list))

    menu = open_group_create_menu_via_no_mouse(group_list)
    if menu is None:
        return "ERROR", "내 프로필 기준 사용자그룹 생성 메뉴 열기 실패"

    ok = select_menu_item_by_keyboard(menu, ["사용자그룹 생성", "사용자그룹 추가"])
    if not ok:
        close_context_menu_only()
        return "ERROR", "사용자그룹 생성/추가 메뉴 선택 실패"

    time.sleep(CFG.MENU_OPEN_WAIT)
    result = type_group_name_and_confirm(AUTO_GROUP_NAME, main_hwnd=main_hwnd)

    if result == "DUPLICATE":
        return "DUPLICATE", "동일한 이름이 존재합니다."

    if result != "OK":
        return "ERROR", "사용자그룹명 입력/생성 실패"

    print(f"{AUTO_GROUP_NAME} 사용자그룹 생성 완료 추정")
    return "OK", f"{AUTO_GROUP_NAME} 생성 완료"


# ----------------------------------------------------------------------
# 검색 / 검증
# ----------------------------------------------------------------------

def click_search_tab(root):
    buttons = find_controls(
        root,
        control_type="Button",
        name="검색",
        auto_id="1018",
    )

    if not buttons:
        print("검색 탭을 찾지 못했습니다.")
        return False

    return invoke_control(buttons[0], "상단 검색 탭")


def choose_filter_child_edit(root):
    edits = find_controls(root, control_type="Edit", class_name="UltariChildEdit")
    small = [e for e in edits if rect_width(e) <= 120]

    if small:
        small.sort(key=lambda e: rect_width(e))
        return small[0]

    return None


def choose_main_search_edit(root):
    edits = find_controls(root, control_type="Edit", class_name="UltariChildEdit")

    if edits:
        edits.sort(key=lambda e: rect_width(e), reverse=True)
        return edits[0]

    edits = find_controls(root, control_type="Edit")

    if edits:
        edits.sort(key=lambda e: rect_width(e), reverse=True)
        return edits[0]

    return None


def set_search_filter_all_no_mouse(root):
    print("\n" + "=" * 80)
    print("[검색 필터 전체 설정 - 무좌표]")
    print("=" * 80)

    filter_edit = choose_filter_child_edit(root)

    if filter_edit is None:
        print("필터 입력칸을 찾지 못했습니다.")
        return False

    try:
        focus_and_replace_text(filter_edit, "전체", "검색 필터")
        keyboard.send_keys("{ENTER}")
        time.sleep(0.18)
        print("검색 필터 전체 설정 완료")
        return True
    except Exception as e:
        print(f"검색 필터 전체 설정 실패: {e}")
        return False


def click_body_search_button(root):
    buttons = find_controls(root, control_type="Button", name="검색")
    body_buttons = [b for b in buttons if info_of(b)["auto_id"] != "1018"]

    if body_buttons:
        body_buttons.sort(key=lambda b: rect_top(b), reverse=True)
        return invoke_control(body_buttons[0], "본문 검색 버튼")

    print("본문 검색 버튼을 찾지 못해 Enter 실행")
    keyboard.send_keys("{ENTER}")
    time.sleep(0.45)
    return True


def click_result_research_button(root):
    buttons = find_controls(
        root,
        control_type="Button",
        name="결과 내 재검색",
        auto_id="1016",
    )

    if not buttons:
        buttons = find_controls(root, control_type="Button", name="결과 내 재검색")

    if not buttons:
        print("결과 내 재검색 버튼을 찾지 못했습니다.")
        return False

    return invoke_control(buttons[0], "결과 내 재검색")


def looks_like_result_text(text):
    text = compact_text(text)
    return ("이 름:" in text and "부 서:" in text) or ("이름:" in text and "부서:" in text)


def collect_result_candidates(root, main_hwnd):
    print("\n" + "=" * 80)
    print("[검색 결과 리스트 후보 수집]")
    print("=" * 80)

    candidates = {}

    try:
        controls = root.descendants()
    except Exception:
        controls = []

    for ctrl in controls:
        try:
            d = info_of(ctrl)

            if not is_visible_rect(ctrl):
                continue

            if d["class_name"] != "UltariListScrollView":
                continue

            hwnd = d["handle"]
            if not hwnd:
                continue

            pieces = []

            if d["name"]:
                pieces.append(d["name"])

            try:
                wt = safe_text(ctrl.window_text())
                if wt:
                    pieces.append(wt)
            except Exception:
                pass

            try:
                for t in ctrl.texts():
                    t = safe_text(t)
                    if t:
                        pieces.append(t)
            except Exception:
                pass

            try:
                legacy = ctrl.legacy_properties()
                if legacy:
                    for _, v in legacy.items():
                        v = safe_text(v)
                        if v:
                            pieces.append(v)
            except Exception:
                pass

            combined = compact_text(" ".join(pieces))

            candidates[hwnd] = {
                "hwnd": hwnd,
                "uia": ctrl,
                "source": "uia",
                "text": combined,
                "info": d,
            }

        except Exception:
            pass

    def enum_child(child_hwnd, _):
        try:
            class_name = safe_text(win32gui.GetClassName(child_hwnd))

            if class_name != "UltariListScrollView":
                return True

            title = safe_text(win32gui.GetWindowText(child_hwnd))
            wmtext = safe_text(wm_gettext(child_hwnd))
            rect = win32gui.GetWindowRect(child_hwnd)

            combined = compact_text(" ".join([title, wmtext]))

            if child_hwnd not in candidates:
                candidates[child_hwnd] = {
                    "hwnd": child_hwnd,
                    "uia": None,
                    "source": "win32",
                    "text": combined,
                    "info": {
                        "name": title,
                        "control_type": "",
                        "class_name": class_name,
                        "auto_id": "",
                        "handle": child_hwnd,
                        "rect": str(rect),
                    },
                }
            else:
                if combined and combined not in candidates[child_hwnd]["text"]:
                    candidates[child_hwnd]["text"] = compact_text(
                        candidates[child_hwnd]["text"] + " " + combined
                    )

        except Exception:
            pass

        return True

    try:
        win32gui.EnumChildWindows(main_hwnd, enum_child, None)
    except Exception as e:
        print(f"EnumChildWindows 실패: {e}")

    rows = list(candidates.values())

    if CFG.VERBOSE_RESULT_SCAN:
        for idx, c in enumerate(rows):
            flag = "RESULT_TEXT" if looks_like_result_text(c["text"]) else ""
            print(
                f"{idx:03d} | {flag} | hwnd={c['hwnd']} | "
                f"text={c['text']} | info={c['info']}"
            )
    else:
        print(f"검색 결과 리스트 후보 수: {len(rows)}")

    result_text_candidates = [c for c in rows if looks_like_result_text(c["text"])]

    if result_text_candidates:
        return result_text_candidates[0]

    generic = [
        c for c in rows
        if "검색 결과 리스트" in c["text"] or c["info"].get("name") == "검색 결과 리스트"
    ]

    if generic:
        return generic[0]

    return None


def reacquire_uia_from_hwnd(hwnd):
    try:
        return Desktop(backend="uia").window(handle=hwnd).wrapper_object()
    except Exception:
        return None


def read_result_text(hwnd, uia_ctrl=None):
    pieces = []

    try:
        title = safe_text(win32gui.GetWindowText(hwnd))
        if title:
            pieces.append(title)
    except Exception:
        pass

    try:
        text = safe_text(wm_gettext(hwnd))
        if text:
            pieces.append(text)
    except Exception:
        pass

    ctrl = uia_ctrl or reacquire_uia_from_hwnd(hwnd)

    if ctrl is not None:
        try:
            d = info_of(ctrl)
            if d["name"]:
                pieces.append(d["name"])
        except Exception:
            pass

        try:
            wt = safe_text(ctrl.window_text())
            if wt:
                pieces.append(wt)
        except Exception:
            pass

        try:
            for t in ctrl.texts():
                t = safe_text(t)
                if t:
                    pieces.append(t)
        except Exception:
            pass

        try:
            legacy = ctrl.legacy_properties()
            if legacy:
                for _, v in legacy.items():
                    v = safe_text(v)
                    if v:
                        pieces.append(v)
        except Exception:
            pass

    unique = []

    for p in pieces:
        p = compact_text(p)
        if p and p not in unique:
            unique.append(p)

    for p in unique:
        if looks_like_result_text(p):
            return p

    return unique[0] if unique else ""



def focus_result(hwnd, uia_ctrl=None):
    print("\n[검색 결과 리스트 포커스]")
    try:
        if uia_ctrl is None:
            uia_ctrl = reacquire_uia_from_hwnd(hwnd)
        if uia_ctrl is not None:
            uia_ctrl.set_focus()
            time.sleep(0.05)
            return True
    except Exception as e:
        print(f"UIA set_focus 실패: {e}")

    try:
        win32gui.SetFocus(hwnd)
        time.sleep(0.05)
        return True
    except Exception as e:
        print(f"Win32 SetFocus 실패: {e}")

    return False

def parse_result_text(text):
    text = compact_text(text)

    name = ""
    dept = ""
    position = ""

    m_name = re.search(r"이\s*름\s*:\s*(.*?)\s+부\s*서\s*:", text)
    if m_name:
        name = safe_text(m_name.group(1))

    m_dept = re.search(r"부\s*서\s*:\s*(.*?)\s+직\s*위\s*:", text)
    if m_dept:
        dept = safe_text(m_dept.group(1))

    m_pos = re.search(
        r"직\s*위\s*:\s*(.*?)(?:\s+일반전화|\s+휴대전화|\s+이메일주소|\s+담당업무|\s+\d{2,4}-|\s+오프라인|\s+온라인|\s+자리비움|$)",
        text,
    )
    if m_pos:
        position = safe_text(m_pos.group(1))

    return {
        "raw": text,
        "name": name,
        "dept": dept,
        "position": position,
    }


def name_matches(target_name, parsed):
    if not target_name:
        return False

    target = normalize(target_name)
    actual = normalize(parsed.get("name", ""))

    return bool(actual and (target == actual or target in actual))


def dept_matches(dept_keyword, parsed):
    if not dept_keyword:
        return False

    target = normalize(dept_keyword)
    dept = normalize(parsed.get("dept", ""))
    raw = normalize(parsed.get("raw", ""))

    return bool((dept and target in dept) or target in raw)


def get_first_two_results(result_hwnd, result_uia):
    rows = []

    if not focus_result(result_hwnd, result_uia):
        return rows

    print("HOME 전송")
    keyboard.send_keys("{HOME}")
    time.sleep(0.05)

    first = compact_text(read_result_text(result_hwnd, result_uia))
    print(f"첫 번째 결과: {first}")
    rows.append(first)

    print("DOWN 전송")
    keyboard.send_keys("{DOWN}")
    time.sleep(0.05)

    second = compact_text(read_result_text(result_hwnd, result_uia))
    print(f"두 번째 결과 후보: {second}")
    rows.append(second)

    try:
        keyboard.send_keys("{HOME}")
        time.sleep(0.07)
    except Exception:
        pass

    return rows


def judge_target(target_name, dept_keyword, rows):
    first = rows[0] if len(rows) >= 1 else ""
    second = rows[1] if len(rows) >= 2 else ""

    first_ok = looks_like_result_text(first)
    second_ok = looks_like_result_text(second)

    if not first_ok:
        return {
            "status": "실패",
            "reason": "검색 결과 없음 또는 결과 읽기 실패",
            "first": parse_result_text(first),
            "second": parse_result_text(second),
        }

    if second_ok and compact_text(second) != compact_text(first):
        return {
            "status": "실패",
            "reason": "이름+부서 재검색 후 복수 결과",
            "first": parse_result_text(first),
            "second": parse_result_text(second),
        }

    parsed = parse_result_text(first)

    if not name_matches(target_name, parsed):
        return {
            "status": "실패",
            "reason": "대상자 이름 불일치",
            "first": parsed,
            "second": parse_result_text(second),
        }

    if not dept_matches(dept_keyword, parsed):
        return {
            "status": "실패",
            "reason": "부서/기관명 불일치",
            "first": parsed,
            "second": parse_result_text(second),
        }

    return {
        "status": "성공",
        "reason": "단일 결과이며 이름·부서 일치",
        "first": parsed,
        "second": parse_result_text(second),
    }


def verify_target(main_hwnd, target_name, first_keyword, dept_keyword):
    root = get_root(main_hwnd)

    if not click_search_tab(root):
        return None, {"status": "실패", "reason": "검색 탭 클릭 실패"}

    time.sleep(CFG.MENU_OPEN_WAIT)
    root = get_root(main_hwnd)

    filter_ok = set_search_filter_all_no_mouse(root)
    print(f"필터 전체 설정 결과: {filter_ok}")

    time.sleep(0.18)
    root = get_root(main_hwnd)

    search_edit = choose_main_search_edit(root)
    if search_edit is None:
        return None, {"status": "실패", "reason": "검색 입력칸 없음"}

    focus_and_replace_text(search_edit, first_keyword, "1차 검색어")

    root = get_root(main_hwnd)

    if not click_body_search_button(root):
        return None, {"status": "실패", "reason": "검색 버튼 클릭 실패"}

    print("1차 검색 결과 대기")
    time.sleep(CFG.SEARCH_WAIT)

    root = get_root(main_hwnd)
    search_edit = choose_main_search_edit(root)
    if search_edit is None:
        return None, {"status": "실패", "reason": "결과 내 재검색 입력칸 없음"}

    focus_and_replace_text(search_edit, dept_keyword, "결과 내 재검색어")

    root = get_root(main_hwnd)

    if not click_result_research_button(root):
        return None, {"status": "실패", "reason": "결과 내 재검색 실패"}

    print("결과 내 재검색 결과 대기")
    time.sleep(CFG.SEARCH_WAIT)

    root = get_root(main_hwnd)
    chosen = collect_result_candidates(root, main_hwnd)

    if chosen is None:
        return None, {"status": "실패", "reason": "검색 결과 리스트 없음"}

    result_hwnd = chosen["hwnd"]
    result_uia = chosen["uia"] or reacquire_uia_from_hwnd(result_hwnd)

    rows = get_first_two_results(result_hwnd, result_uia)
    judgment = judge_target(target_name, dept_keyword, rows)

    return {
        "hwnd": result_hwnd,
        "uia": result_uia,
        "rows": rows,
        "parsed": judgment.get("first", {}),
    }, judgment


# ----------------------------------------------------------------------
# 사용자그룹 선택창 처리
# ----------------------------------------------------------------------

def get_select_frame_handles():
    handles = set()

    for win in Desktop(backend="uia").windows():
        try:
            controls = [win] + win.descendants()
        except Exception:
            controls = [win]

        for ctrl in controls:
            try:
                d = info_of(ctrl)

                if d["name"] == "JBEdu Messenger+ Frame Window" and d["class_name"] == "@Messenger7_Wnd":
                    h = d["handle"]

                    names = []
                    try:
                        for child in ctrl.descendants():
                            cd = info_of(child)
                            if cd["name"]:
                                names.append(cd["name"])
                    except Exception:
                        pass

                    if h and "확인" in names and "취소" in names:
                        handles.add(h)
            except Exception:
                pass

    return handles


def find_select_frame(existing_handles=None, timeout=8):
    if existing_handles is None:
        existing_handles = set()

    end = time.time() + timeout
    fallback = None

    while time.time() < end:
        candidates = []

        for win in Desktop(backend="uia").windows():
            try:
                controls = [win] + win.descendants()
            except Exception:
                controls = [win]

            for ctrl in controls:
                try:
                    d = info_of(ctrl)

                    if d["name"] == "JBEdu Messenger+ Frame Window" and d["class_name"] == "@Messenger7_Wnd":
                        candidates.append(ctrl)
                except Exception:
                    pass

        for cand in candidates:
            try:
                d = info_of(cand)
                h = d["handle"]

                names = []
                for child in cand.descendants():
                    cd = info_of(child)
                    if cd["name"]:
                        names.append(cd["name"])

                if "확인" in names and "취소" in names:
                    fallback = cand

                    if h not in existing_handles:
                        return cand
            except Exception:
                pass

        time.sleep(0.10)

    return fallback


def reacquire_frame_from_hwnd(hwnd):
    try:
        if hwnd and win32gui.IsWindow(hwnd):
            return Desktop(backend="uia").window(handle=hwnd).wrapper_object()
    except Exception as e:
        print(f"선택 프레임 재획득 실패: {e}")
    return None


def is_hwnd_alive_visible(hwnd):
    try:
        return bool(hwnd and win32gui.IsWindow(hwnd) and win32gui.IsWindowVisible(hwnd))
    except Exception:
        return False


def find_in_descendants(root, control_type=None, name=None, auto_id=None, class_name=None):
    try:
        controls = root.descendants()
    except Exception:
        controls = []

    for ctrl in controls:
        try:
            d = info_of(ctrl)

            if not is_visible_rect(ctrl):
                continue

            if control_type is not None and d["control_type"] != control_type:
                continue

            if name is not None and d["name"] != name:
                continue

            if auto_id is not None and d["auto_id"] != str(auto_id):
                continue

            if class_name is not None and d["class_name"] != class_name:
                continue

            return ctrl
        except Exception:
            pass

    return None


def collect_select_popup_snapshot(except_hwnd=None):
    """
    사용자그룹 선택창에서 확인을 눌렀을 때 뜨는 안내/오류 팝업을 찾기 위한 진단용 수집.
    메신저 팝업은 문구가 UIA에 노출되지 않는 경우가 있어 크기/버튼 구성도 같이 본다.
    """
    rows = []

    for win in Desktop(backend="uia").windows():
        try:
            d = info_of(win)
            h = d.get("handle")
            if not h or h == except_hwnd:
                continue
            if d.get("class_name") != "@Messenger7_Wnd":
                continue
            if not is_hwnd_alive_visible(h):
                continue

            rect = win32gui.GetWindowRect(h)
            width = rect[2] - rect[0]
            height = rect[3] - rect[1]

            child_names = []
            button_names = []
            edit_count = 0
            try:
                for child in win.descendants():
                    cd = info_of(child)
                    if cd["name"]:
                        child_names.append(cd["name"])
                    if cd["control_type"] == "Button" and cd["name"]:
                        button_names.append(cd["name"])
                    if cd["control_type"] == "Edit":
                        edit_count += 1
            except Exception:
                pass

            texts = collect_texts_from_window_handle(h)
            combined = compact_text(" ".join(texts + child_names + button_names))

            rows.append({
                "hwnd": h,
                "info": d,
                "rect": rect,
                "width": width,
                "height": height,
                "button_names": button_names,
                "edit_count": edit_count,
                "text": combined,
            })
        except Exception:
            pass

    return rows


def close_possible_select_required_popup(frame_hwnd):
    """
    '추가할 사용자그룹을 선택해 주세요' 같은 안내 팝업이 떠 있으면 닫는다.
    문구가 안 읽히는 경우가 있어 작은 @Messenger7_Wnd 팝업도 안내창으로 간주한다.
    """
    rows = collect_select_popup_snapshot(except_hwnd=frame_hwnd)

    for row in rows:
        text = row.get("text", "")
        width = row.get("width", 9999)
        height = row.get("height", 9999)
        button_names = row.get("button_names", [])
        edit_count = row.get("edit_count", 0)
        hwnd = row.get("hwnd")

        looks_like_message = (
            "추가할 사용자그룹" in text
            or "선택" in text
            or "확인" in button_names
            or (width <= 360 and height <= 180 and edit_count == 0)
        )

        if not looks_like_message:
            continue

        print("\n[선택 필요/안내 팝업 감지 - 닫기 시도]")
        print(hwnd_info(hwnd))
        print(f"button_names={button_names}, edit_count={edit_count}, size=({width}x{height})")
        if text:
            print(text[:1000])

        try:
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.05)
        except Exception as e:
            print(f"안내 팝업 foreground 실패, 계속 진행: {e}")

        try:
            keyboard.send_keys("{ENTER}")
            time.sleep(0.45)
            print("안내 팝업 Enter 닫기 완료")
            return True
        except Exception as e:
            print(f"안내 팝업 닫기 실패: {e}")

    return False


def invoke_confirm_and_check(frame_hwnd, confirm_btn, label):
    """
    확인 버튼을 누른 뒤 선택창이 닫혔는지 확인한다.
    선택이 안 된 경우 안내창이 뜨고 선택창이 남아 있으므로 실패로 반환한다.
    """
    if confirm_btn is None:
        print("확인 버튼이 없어 확인 실행 불가")
        return False

    if not invoke_control(confirm_btn, label):
        print("확인 버튼 실행 자체가 실패했습니다.")
        return False

    time.sleep(CFG.AFTER_CONFIRM_WAIT)

    if not is_hwnd_alive_visible(frame_hwnd):
        print("사용자그룹 선택창이 닫혔습니다. 확인 성공으로 판단합니다.")
        return True

    closed_popup = close_possible_select_required_popup(frame_hwnd)
    if closed_popup:
        print("선택 필요 안내창이 떠서 닫았습니다. 이번 확인은 실패로 판단합니다.")
    else:
        print("확인 후에도 사용자그룹 선택창이 열려 있습니다. 이번 확인은 실패로 판단합니다.")

    return False


def send_win32_key_to_hwnd(hwnd, vk, label=""):
    """
    키보드 포커스가 UIA set_focus만으로 안 들어가는 경우를 대비해
    해당 UltariListScrollView 핸들에 직접 키 메시지를 보낸다.
    마우스 좌표 클릭은 사용하지 않는다.
    """
    try:
        win32gui.SendMessage(hwnd, win32con.WM_KEYDOWN, vk, 0)
        time.sleep(0.06)
        win32gui.SendMessage(hwnd, win32con.WM_KEYUP, vk, 0)
        time.sleep(0.05)
        if label:
            print(f"Win32 키 전송: {label}")
        return True
    except Exception as e:
        print(f"Win32 키 전송 실패({label}): {e}")
        return False


def key_token_to_vk(token):
    return {
        "LEFT": win32con.VK_LEFT,
        "RIGHT": win32con.VK_RIGHT,
        "UP": win32con.VK_UP,
        "DOWN": win32con.VK_DOWN,
        "SPACE": win32con.VK_SPACE,
        "ENTER": win32con.VK_RETURN,
        "PGUP": win32con.VK_PRIOR,
        "PGDN": win32con.VK_NEXT,
        "TAB": win32con.VK_TAB,
    }.get(token)


def focus_tree_and_send_sequence(group_tree, sequence, label, use_win32=False):
    print(f"\n[첫 번째 그룹 키보드 선택 시도] {label}")
    print(info_of(group_tree))

    try:
        group_tree.set_focus()
        time.sleep(0.08)

        # 일부 PC에서는 UIA set_focus 후에도 실제 리스트 항목에는 키가 안 들어간다.
        # 그래서 가능한 경우 Win32 SetFocus도 한 번 더 시도한다.
        tree_hwnd = info_of(group_tree).get("handle")
        if tree_hwnd:
            try:
                win32gui.SetFocus(tree_hwnd)
                time.sleep(0.05)
                print(f"Win32 SetFocus 성공: {tree_hwnd}")
            except Exception as e:
                print(f"Win32 SetFocus 실패, UIA 포커스로 계속 진행: {e}")

        if use_win32 and tree_hwnd:
            for token in sequence:
                vk = key_token_to_vk(token)
                if vk is None:
                    continue
                send_win32_key_to_hwnd(tree_hwnd, vk, token)
            print(f"조직도 트리 Win32 키 시퀀스 전송 완료: {' → '.join(sequence)}")
            return True

        # HOME은 사용자그룹 탭과 선택창에서 먹지 않는 경우가 확인되어 사용하지 않는다.
        key_string = "".join("{" + token + "}" for token in sequence)
        if key_string:
            keyboard.send_keys(key_string)
            time.sleep(0.05)
        print(f"조직도 트리 키보드 시퀀스 전송 완료: {' → '.join(sequence) if sequence else '(없음)'}")
        return True
    except Exception as e:
        print(f"조직도 트리 키보드 선택 시도 실패: {e}")
        return False


def try_keyboard_first_group_select_and_confirm(frame_hwnd):
    """
    v46 핵심:
    - 선택창의 조직도 트리에서도 HOME이 먹지 않는 경우가 있어 HOME을 사용하지 않는다.
    - 사용자그룹 탭에서 안정적이었던 LEFT 기준을 내목록 추가 선택창에도 적용한다.
    - 성공 조건은 확인 후 사용자그룹 선택창이 닫히는 것이다.
    - 좌표 클릭은 사용하지 않는다.
    """
    attempts = [
        (["LEFT"], "LEFT만"),
        (["LEFT", "SPACE"], "LEFT 후 SPACE"),
        (["LEFT", "ENTER"], "LEFT 후 ENTER"),
        (["UP", "LEFT", "SPACE"], "UP → LEFT → SPACE"),
        (["PGUP", "LEFT", "SPACE"], "PGUP → LEFT → SPACE"),
        (["RIGHT", "LEFT", "SPACE"], "RIGHT → LEFT → SPACE"),
        (["LEFT", "DOWN", "UP", "SPACE"], "LEFT → DOWN → UP → SPACE"),
        (["LEFT"], "Win32 LEFT만", True),
        (["LEFT", "SPACE"], "Win32 LEFT 후 SPACE", True),
        (["LEFT", "ENTER"], "Win32 LEFT 후 ENTER", True),
    ]

    for attempt in attempts:
        if len(attempt) == 2:
            sequence, label = attempt
            use_win32 = False
        else:
            sequence, label, use_win32 = attempt

        frame = reacquire_frame_from_hwnd(frame_hwnd)
        if frame is None:
            print("선택 프레임이 이미 사라졌습니다. 이전 단계 성공으로 판단합니다.")
            return True

        group_tree = find_in_descendants(
            frame,
            control_type="Pane",
            name="조직도 트리",
            class_name="UltariListScrollView",
        )
        if group_tree is None:
            print("조직도 트리를 찾지 못했습니다. 키보드 방식 중단")
            return False

        confirm_btn = find_in_descendants(
            frame,
            control_type="Button",
            name="확인",
            auto_id="1003",
        )
        if confirm_btn is None:
            print("확인 버튼을 찾지 못했습니다. 키보드 방식 중단")
            return False

        if not focus_tree_and_send_sequence(group_tree, sequence, label, use_win32=use_win32):
            continue

        if invoke_confirm_and_check(frame_hwnd, confirm_btn, f"사용자그룹 선택 확인 - 키보드 방식({label})"):
            print("[키보드 선택 성공] 좌표 클릭 없이 첫 번째 그룹 선택/확인이 완료되었습니다.")
            return True

        print(f"키보드 방식 실패: {label}")

    print("[키보드 선택 실패] 모든 LEFT 계열 키보드 방식에서 선택창이 닫히지 않았습니다.")
    return False


def make_lparam(x, y):
    return ((int(y) & 0xFFFF) << 16) | (int(x) & 0xFFFF)


def find_group_tree_in_frame(frame):
    """
    내목록 추가 창 안의 조직도 트리 컨트롤을 찾는다.
    UIA에서 못 찾을 경우 Win32 자식 창에서도 UltariListScrollView/조직도 트리를 보조 탐색한다.
    """
    group_tree = find_in_descendants(
        frame,
        control_type="Pane",
        name="조직도 트리",
        class_name="UltariListScrollView",
    )
    if group_tree is not None:
        return group_tree

    try:
        frame_hwnd = info_of(frame).get("handle")
    except Exception:
        frame_hwnd = None

    if not frame_hwnd:
        return None

    found = []

    def enum_child(hwnd, _):
        try:
            title = safe_text(win32gui.GetWindowText(hwnd))
            class_name = safe_text(win32gui.GetClassName(hwnd))
            if class_name == "UltariListScrollView" and title == "조직도 트리":
                found.append(hwnd)
        except Exception:
            pass
        return True

    try:
        win32gui.EnumChildWindows(frame_hwnd, enum_child, None)
    except Exception:
        pass

    if found:
        try:
            return Desktop(backend="uia").window(handle=found[0]).wrapper_object()
        except Exception:
            return None

    return None


def send_top_row_select_message_to_tree(tree_hwnd, frame_hwnd=None):
    """
    선택창의 조직도 트리에서 첫 번째 행을 선택한다.

    중요:
    - 실제 마우스 커서를 이동하지 않는다.
    - 화면 절대좌표를 찍지 않는다.
    - 진단 결과 기준으로 UltariListScrollView 내부 첫 번째 행의 client 좌표에
      WM_LBUTTONDOWN/UP 메시지만 보낸다.
    """
    if not tree_hwnd or not win32gui.IsWindow(tree_hwnd):
        print("조직도 트리 hwnd가 유효하지 않습니다.")
        return False

    try:
        left, top, right, bottom = win32gui.GetWindowRect(tree_hwnd)
        width = max(1, int(right - left))
        height = max(1, int(bottom - top))
    except Exception as e:
        print(f"조직도 트리 rect 확인 실패: {e}")
        return False

    # 사용자가 직접 클릭해 진단한 위치는 tree client 기준 대략 x=81, y=18이었다.
    # PC별 글자/아이콘 위치가 조금 달라도 첫 줄 내부가 되도록 x 후보를 몇 개 둔다.
    y_candidates = [18, 16, 22, 28]
    x_candidates = [
        min(max(80, int(width * 0.28)), width - 8),
        min(max(45, int(width * 0.16)), width - 8),
        min(max(110, int(width * 0.38)), width - 8),
    ]

    points = []
    for y in y_candidates:
        if y >= height:
            continue
        for x in x_candidates:
            if x >= width:
                continue
            p = (int(x), int(y))
            if p not in points:
                points.append(p)

    print("\n[첫 번째 그룹 선택 - 조직도 트리 첫 행 메시지 선택]")
    print(hwnd_info(tree_hwnd))
    print(f"tree client size=({width}x{height})")
    print(f"시도 좌표 후보(client): {points}")
    print("※ 실제 마우스 커서는 이동하지 않고, 조직도 트리 hwnd에 WM_LBUTTONDOWN/UP만 전송합니다.")

    try:
        if frame_hwnd and win32gui.IsWindow(frame_hwnd):
            win32gui.SetForegroundWindow(frame_hwnd)
            time.sleep(0.05)
    except Exception as e:
        print(f"선택창 foreground 실패, 계속 진행: {e}")

    try:
        win32gui.SetFocus(tree_hwnd)
        time.sleep(0.05)
    except Exception as e:
        print(f"조직도 트리 SetFocus 실패, 계속 진행: {e}")

    for idx, (x, y) in enumerate(points, start=1):
        try:
            lp = make_lparam(x, y)
            print(f"첫 행 선택 메시지 전송 {idx}: client=({x}, {y})")
            win32gui.SendMessage(tree_hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lp)
            time.sleep(0.08)
            win32gui.SendMessage(tree_hwnd, win32con.WM_LBUTTONUP, 0, lp)
            time.sleep(0.08)
            return True
        except Exception as e:
            print(f"첫 행 선택 메시지 전송 실패 {idx}: {e}")

    return False


def fallback_click_first_group_and_confirm(frame_hwnd):
    """
    v46:
    첫 번째 그룹명을 찾지 않고, 조직도 트리의 첫 번째 행을 선택한 뒤 확인한다.
    실제 마우스 클릭은 사용하지 않고 Win32 메시지를 보낸다.
    """
    print("\n[첫 번째 그룹 선택 - 첫 행 직접 선택]")
    print("그룹명 검색은 하지 않습니다. 내목록 추가 창에서 가장 위에 있는 그룹을 선택합니다.")

    frame = reacquire_frame_from_hwnd(frame_hwnd)
    if frame is None:
        print("선택 프레임이 이미 사라졌습니다. 이전 단계 성공으로 판단합니다.")
        return True

    group_tree = find_group_tree_in_frame(frame)
    if group_tree is None:
        print("조직도 트리를 찾지 못했습니다. 첫 행 선택 실패")
        return False

    tree_hwnd = info_of(group_tree).get("handle")
    if not send_top_row_select_message_to_tree(tree_hwnd, frame_hwnd=frame_hwnd):
        print("조직도 트리 첫 행 선택 메시지 전송 실패")
        return False

    frame = reacquire_frame_from_hwnd(frame_hwnd)
    if frame is None:
        print("선택 프레임이 첫 행 선택 후 사라졌습니다. 성공으로 판단합니다.")
        return True

    confirm_btn = find_in_descendants(
        frame,
        control_type="Button",
        name="확인",
        auto_id="1003",
    )
    if confirm_btn is None:
        print("확인 버튼을 찾지 못했습니다. 첫 행 선택 실패")
        return False

    if invoke_confirm_and_check(frame_hwnd, confirm_btn, "사용자그룹 선택 확인 - 첫 행 선택 메시지"):
        print("[첫 행 선택 성공] 조직도 트리 첫 번째 그룹 선택 후 확인 완료")
        return True

    print("첫 행 선택 메시지 후에도 선택창이 닫히지 않았습니다.")
    return False


def select_first_group_and_confirm(frame):
    print("\n" + "=" * 80)
    print("[첫 번째 사용자그룹 선택 후 확인 - v46 바로 첫 행 선택]")
    print("=" * 80)

    print("선택 프레임:")
    print(info_of(frame))

    frame_hwnd = info_of(frame).get("handle")
    if not frame_hwnd:
        print("선택 프레임 핸들을 찾지 못했습니다.")
        return False

    print("키보드 선택 시도는 하지 않습니다.")
    print("내목록 추가 창에서 조직도 트리의 가장 위 첫 번째 행을 바로 선택합니다.")
    return fallback_click_first_group_and_confirm(frame_hwnd)

def add_verified_target_to_first_group(verify_result, expected_name):
    result_hwnd = verify_result["hwnd"]
    result_uia = verify_result["uia"]

    existing_frames = get_select_frame_handles()

    menu = open_context_menu_with_single_vk_apps(result_hwnd, result_uia, expected_name)

    if menu is None:
        return False, "컨텍스트 메뉴 열기 실패"

    menu_ok, first_line = validate_menu_first_item(menu, expected_name)

    if not menu_ok:
        close_context_menu_only()
        return False, f"컨텍스트 메뉴 첫 줄 대상자 불일치: {first_line}"

    ok = select_menu_item_by_keyboard(menu, "내목록 추가")

    if not ok:
        close_context_menu_only()
        return False, "내목록 추가 메뉴 키보드 선택 실패"

    print("사용자그룹 선택창 대기")
    frame = find_select_frame(existing_handles=existing_frames, timeout=8)

    if frame is None:
        return False, "사용자그룹 선택창 찾기 실패"

    if not select_first_group_and_confirm(frame):
        return False, "첫 번째 그룹 선택/확인 실패"

    return True, "사용자그룹 추가 완료 추정"


def validate_menu_first_item(menu, expected_name):
    rows = get_menu_items(menu)

    first_name = ""
    for row in rows:
        name = row["info"]["name"]
        if name:
            first_name = name
            break

    print(f"\n컨텍스트 메뉴 첫 줄: {first_name}")

    if not expected_name:
        return True, first_name

    ok = normalize(expected_name) in normalize(first_name)
    return ok, first_name



def open_context_menu_with_single_vk_apps(result_hwnd, result_uia, expected_name):
    print("\n[검색결과 컨텍스트 메뉴 열기]")
    close_context_menu_only()

    if not focus_result(result_hwnd, result_uia):
        print("결과 리스트 포커스 실패")
        return None

    # 검증 단계에서 이미 첫 번째 결과가 대상자임을 확인했다.
    # 그래도 HOME만 보내 현재 선택을 첫 행으로 맞춘다.
    keyboard.send_keys("{HOME}")
    time.sleep(0.05)

    try:
        keyboard.send_keys("{VK_APPS}")
        time.sleep(CFG.MENU_OPEN_WAIT)
    except Exception as e:
        print(f"VK_APPS 전송 실패: {e}")
        return None

    menu = find_context_menu()
    if menu is None:
        print("컨텍스트 메뉴를 찾지 못했습니다.")
        return None

    print("컨텍스트 메뉴 열림")
    return menu



# ----------------------------------------------------------------------
# v48 FAST TWO-CHECK OVERRIDES
# ----------------------------------------------------------------------
# 병목 1: 검색 결과 개수/복수 결과 검증을 줄인다.
# - 전체 UIA descendants 스캔을 하지 않고 Win32 자식창만 빠르게 훑는다.
# - 결과 내 재검색 후 첫 번째 결과만 이름/부서 일치 검증한다.
# - 100명 일괄 처리에서는 이 방식이 훨씬 빠르다.
#
# 병목 2: 내목록 추가 선택창 대기/첫 번째 그룹 선택을 줄인다.
# - 선택창 탐색은 Win32 top-level/window child 기준으로 먼저 찾는다.
# - 조직도 트리와 확인 버튼도 Win32 핸들로 바로 찾는다.
# - 첫 번째 그룹 선택은 v44 방식 그대로 유지하되, 후보 좌표 반복/상세 로그를 줄인다.

# 빠른 경로에서는 고정 대기 대신 짧은 폴링을 사용한다.
CFG.SEARCH_WAIT = 0.35
CFG.MENU_OPEN_WAIT = 0.20
CFG.MENU_ACTION_WAIT = 0.16
CFG.AFTER_FIRST_GROUP_CLICK_DELAY = 0.08
CFG.AFTER_CONFIRM_WAIT = 0.12
CFG.INPUT_WAIT = 0.08
CFG.SHORT_WAIT = 0.035


def win32_child_windows(parent_hwnd):
    rows = []

    def cb(hwnd, _):
        try:
            rows.append(hwnd)
        except Exception:
            pass
        return True

    try:
        win32gui.EnumChildWindows(parent_hwnd, cb, None)
    except Exception:
        pass
    return rows


def win32_top_windows():
    rows = []

    def cb(hwnd, _):
        try:
            rows.append(hwnd)
        except Exception:
            pass
        return True

    try:
        win32gui.EnumWindows(cb, None)
    except Exception:
        pass
    return rows


def is_visible_nonzero_hwnd(hwnd):
    try:
        if not win32gui.IsWindow(hwnd) or not win32gui.IsWindowVisible(hwnd):
            return False
        l, t, r, b = win32gui.GetWindowRect(hwnd)
        return (r - l) > 5 and (b - t) > 5
    except Exception:
        return False


def get_hwnd_text_fast(hwnd):
    parts = []
    try:
        title = safe_text(win32gui.GetWindowText(hwnd))
        if title:
            parts.append(title)
    except Exception:
        pass
    try:
        wt = safe_text(wm_gettext(hwnd))
        if wt and wt not in parts:
            parts.append(wt)
    except Exception:
        pass
    return compact_text(" ".join(parts))


def find_result_list_fast(main_hwnd, must_contain=None):
    """검색 결과 리스트를 Win32 핸들 기준으로 빠르게 찾는다."""
    must_norm = normalize(must_contain or "")
    candidates = []

    for hwnd in win32_child_windows(main_hwnd):
        try:
            if not is_visible_nonzero_hwnd(hwnd):
                continue
            class_name = safe_text(win32gui.GetClassName(hwnd))
            if class_name != "UltariListScrollView":
                continue
            rect = win32gui.GetWindowRect(hwnd)
            width = rect[2] - rect[0]
            height = rect[3] - rect[1]
            if width < 100 or height < 100:
                continue
            text = get_hwnd_text_fast(hwnd)
            score = 0
            if looks_like_result_text(text):
                score += 100
            if "검색 결과 리스트" in text:
                score += 40
            if must_norm and must_norm in normalize(text):
                score += 30
            # 검색 결과 영역은 보통 검색창 아래의 큰 리스트다.
            area = width * height
            candidates.append((score, area, hwnd, text, rect))
        except Exception:
            pass

    if not candidates:
        return None

    candidates.sort(key=lambda x: (x[0], x[1]), reverse=True)

    # must_contain이 있으면 그 문구를 포함하는 결과를 우선한다.
    if must_norm:
        for score, area, hwnd, text, rect in candidates:
            if looks_like_result_text(text) and must_norm in normalize(text):
                return {"hwnd": hwnd, "uia": None, "text": text, "info": {"name": text, "handle": hwnd, "rect": str(rect)}}

    # 결과 텍스트가 있는 리스트를 우선한다.
    for score, area, hwnd, text, rect in candidates:
        if looks_like_result_text(text):
            return {"hwnd": hwnd, "uia": None, "text": text, "info": {"name": text, "handle": hwnd, "rect": str(rect)}}

    # 없으면 '검색 결과 리스트'라고 보이는 큰 리스트를 반환한다.
    for score, area, hwnd, text, rect in candidates:
        if "검색 결과 리스트" in text:
            return {"hwnd": hwnd, "uia": None, "text": text, "info": {"name": text, "handle": hwnd, "rect": str(rect)}}

    return None


def wait_result_list_fast(main_hwnd, must_contain=None, timeout=1.05):
    """고정 sleep 대신 결과 텍스트가 보일 때까지만 짧게 기다린다."""
    end = time.time() + timeout
    last = None
    while time.time() < end:
        chosen = find_result_list_fast(main_hwnd, must_contain=must_contain)
        if chosen is not None:
            last = chosen
            text = safe_text(chosen.get("text"))
            if looks_like_result_text(text):
                if not must_contain or normalize(must_contain) in normalize(text):
                    return chosen
            elif not must_contain:
                return chosen
        time.sleep(0.045)
    return last


def read_result_text_fast(hwnd):
    """선택된 검색 결과 텍스트를 Win32 우선으로 읽는다."""
    text = get_hwnd_text_fast(hwnd)
    if looks_like_result_text(text):
        return text

    # 꼭 필요할 때만 UIA를 한 번 사용한다.
    try:
        ctrl = Desktop(backend="uia").window(handle=hwnd).wrapper_object()
        pieces = []
        d = info_of(ctrl)
        if d.get("name"):
            pieces.append(d.get("name"))
        try:
            wt = safe_text(ctrl.window_text())
            if wt:
                pieces.append(wt)
        except Exception:
            pass
        text = compact_text(" ".join(pieces))
        if looks_like_result_text(text):
            return text
    except Exception:
        pass

    return text


def focus_result_fast(hwnd):
    try:
        ctrl = Desktop(backend="uia").window(handle=hwnd).wrapper_object()
        ctrl.set_focus()
        time.sleep(0.035)
        return True
    except Exception:
        pass
    try:
        win32gui.SetFocus(hwnd)
        time.sleep(0.035)
        return True
    except Exception:
        return False


def judge_target_first_only(target_name, dept_keyword, first_text):
    first = compact_text(first_text)
    if not looks_like_result_text(first):
        return {
            "status": "실패",
            "reason": "검색 결과 없음 또는 결과 읽기 실패",
            "first": parse_result_text(first),
            "second": {},
        }

    parsed = parse_result_text(first)
    if not name_matches(target_name, parsed):
        return {
            "status": "실패",
            "reason": "대상자 이름 불일치",
            "first": parsed,
            "second": {},
        }

    if not dept_matches(dept_keyword, parsed):
        return {
            "status": "실패",
            "reason": "부서/기관명 불일치",
            "first": parsed,
            "second": {},
        }

    return {
        "status": "성공",
        "reason": "첫 번째 결과 이름·부서 일치(복수 결과 확인 생략 - 빠른 모드)",
        "first": parsed,
        "second": {},
    }


def verify_target(main_hwnd, target_name, first_keyword, dept_keyword):
    """v48 빠른 검증: 첫 번째와 두 번째 결과만 확인해 없음/복수/불일치/성공을 구분."""
    root = get_root(main_hwnd)

    if not click_search_tab(root):
        return None, {"status": "실패", "reason": "검색 탭 클릭 실패"}

    time.sleep(0.16)
    root = get_root(main_hwnd)

    filter_ok = set_search_filter_all_no_mouse(root)
    print(f"필터 전체 설정 결과: {filter_ok}")

    time.sleep(0.08)
    root = get_root(main_hwnd)

    search_edit = choose_main_search_edit(root)
    if search_edit is None:
        return None, {"status": "실패", "reason": "검색 입력칸 없음"}

    focus_and_replace_text(search_edit, first_keyword, "1차 검색어")

    root = get_root(main_hwnd)
    if not click_body_search_button(root):
        return None, {"status": "실패", "reason": "검색 버튼 클릭 실패"}

    print("1차 검색 결과 대기(빠른 폴링)")
    wait_result_list_fast(main_hwnd, must_contain=target_name, timeout=0.95)

    root = get_root(main_hwnd)
    search_edit = choose_main_search_edit(root)
    if search_edit is None:
        return None, {"status": "실패", "reason": "결과 내 재검색 입력칸 없음"}

    focus_and_replace_text(search_edit, dept_keyword, "결과 내 재검색어")

    root = get_root(main_hwnd)
    if not click_result_research_button(root):
        return None, {"status": "실패", "reason": "결과 내 재검색 실패"}

    print("결과 내 재검색 결과 대기(빠른 폴링)")
    chosen = wait_result_list_fast(main_hwnd, must_contain=dept_keyword, timeout=1.05)
    if chosen is None:
        chosen = find_result_list_fast(main_hwnd)

    if chosen is None:
        return None, {"status": "실패", "reason": "검색 결과 리스트 없음"}

    result_hwnd = chosen["hwnd"]

    if not focus_result_fast(result_hwnd):
        return None, {"status": "실패", "reason": "검색 결과 리스트 포커스 실패"}

    # 현재 선택이 첫 행이라고 보는 PC도 있지만, 안전하게 HOME만 한 번 보낸다.
    try:
        keyboard.send_keys("{HOME}")
        time.sleep(0.045)
    except Exception:
        pass

    first_text = read_result_text_fast(result_hwnd)
    print(f"첫 번째 결과 확인: {first_text}")
    judgment = judge_target_first_only(target_name, dept_keyword, first_text)

    return {
        "hwnd": result_hwnd,
        "uia": None,
        "rows": [first_text],
        "parsed": judgment.get("first", {}),
    }, judgment


def get_select_frame_handles():
    """현재 떠 있는 메신저 프레임 창 핸들을 Win32로 빠르게 수집."""
    handles = set()
    for hwnd in win32_top_windows():
        try:
            if not is_visible_nonzero_hwnd(hwnd):
                continue
            title = safe_text(win32gui.GetWindowText(hwnd))
            class_name = safe_text(win32gui.GetClassName(hwnd))
            if title == "JBEdu Messenger+ Frame Window" and class_name == "@Messenger7_Wnd":
                handles.add(hwnd)
        except Exception:
            pass
    return handles


def frame_has_confirm_cancel_fast(hwnd):
    names = []
    for ch in win32_child_windows(hwnd):
        try:
            txt = safe_text(win32gui.GetWindowText(ch))
            if txt:
                names.append(txt)
        except Exception:
            pass
    joined = " ".join(names)
    if "확인" in joined and "취소" in joined:
        return True

    # 일부 PC는 Win32로 버튼 텍스트가 안 잡힐 수 있어 크기로 보조 판단한다.
    try:
        l, t, r, b = win32gui.GetWindowRect(hwnd)
        w, h = r - l, b - t
        if 380 <= w <= 720 and 320 <= h <= 720:
            return True
    except Exception:
        pass
    return False


def find_select_frame(existing_handles=None, timeout=2.0):
    """내목록 추가 선택창을 Win32 우선으로 빠르게 찾는다."""
    if existing_handles is None:
        existing_handles = set()

    end = time.time() + timeout
    fallback_hwnd = None

    while time.time() < end:
        # 새로 foreground가 된 선택창을 먼저 확인한다.
        try:
            fg = win32gui.GetForegroundWindow()
            if fg and fg not in existing_handles and is_visible_nonzero_hwnd(fg):
                title = safe_text(win32gui.GetWindowText(fg))
                class_name = safe_text(win32gui.GetClassName(fg))
                if title == "JBEdu Messenger+ Frame Window" and class_name == "@Messenger7_Wnd":
                    if frame_has_confirm_cancel_fast(fg):
                        return Desktop(backend="uia").window(handle=fg).wrapper_object()
                    fallback_hwnd = fg
        except Exception:
            pass

        for hwnd in win32_top_windows():
            try:
                if hwnd in existing_handles:
                    continue
                if not is_visible_nonzero_hwnd(hwnd):
                    continue
                title = safe_text(win32gui.GetWindowText(hwnd))
                class_name = safe_text(win32gui.GetClassName(hwnd))
                if title != "JBEdu Messenger+ Frame Window" or class_name != "@Messenger7_Wnd":
                    continue
                if frame_has_confirm_cancel_fast(hwnd):
                    return Desktop(backend="uia").window(handle=hwnd).wrapper_object()
                fallback_hwnd = hwnd
            except Exception:
                pass

        time.sleep(0.045)

    if fallback_hwnd:
        try:
            return Desktop(backend="uia").window(handle=fallback_hwnd).wrapper_object()
        except Exception:
            pass
    return None


def find_group_tree_hwnd_fast(frame_hwnd):
    best = None
    best_area = 0
    for hwnd in win32_child_windows(frame_hwnd):
        try:
            if not is_visible_nonzero_hwnd(hwnd):
                continue
            title = safe_text(win32gui.GetWindowText(hwnd))
            class_name = safe_text(win32gui.GetClassName(hwnd))
            if class_name != "UltariListScrollView":
                continue
            if title != "조직도 트리":
                continue
            l, t, r, b = win32gui.GetWindowRect(hwnd)
            area = (r - l) * (b - t)
            if area > best_area:
                best = hwnd
                best_area = area
        except Exception:
            pass
    return best


def find_button_hwnd_by_text_fast(frame_hwnd, button_text):
    for hwnd in win32_child_windows(frame_hwnd):
        try:
            if not is_visible_nonzero_hwnd(hwnd):
                continue
            txt = safe_text(win32gui.GetWindowText(hwnd))
            cls = safe_text(win32gui.GetClassName(hwnd))
            if txt == button_text and (cls == "Button" or "Button" in cls):
                return hwnd
        except Exception:
            pass
    # class명이 다르게 잡히는 경우 텍스트만으로 보조 탐색
    for hwnd in win32_child_windows(frame_hwnd):
        try:
            if is_visible_nonzero_hwnd(hwnd) and safe_text(win32gui.GetWindowText(hwnd)) == button_text:
                return hwnd
        except Exception:
            pass
    return None


def send_top_row_select_message_to_tree(tree_hwnd, frame_hwnd=None):
    """v48 빠른 첫 행 선택: 성공 확인된 첫 후보만 전송하고 상세 후보 반복을 생략."""
    if not tree_hwnd or not win32gui.IsWindow(tree_hwnd):
        print("조직도 트리 hwnd가 유효하지 않습니다.")
        return False

    try:
        l, t, r, b = win32gui.GetWindowRect(tree_hwnd)
        width = max(1, r - l)
        height = max(1, b - t)
    except Exception as e:
        print(f"조직도 트리 rect 확인 실패: {e}")
        return False

    x = min(max(80, int(width * 0.28)), width - 8)
    y = 18 if height > 24 else max(1, height // 2)

    try:
        if frame_hwnd and win32gui.IsWindow(frame_hwnd):
            try:
                win32gui.SetForegroundWindow(frame_hwnd)
            except Exception:
                pass
        lp = make_lparam(x, y)
        win32gui.SendMessage(tree_hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lp)
        time.sleep(0.035)
        win32gui.SendMessage(tree_hwnd, win32con.WM_LBUTTONUP, 0, lp)
        time.sleep(0.055)
        print(f"첫 번째 그룹 첫 행 선택 메시지 전송 완료: client=({x},{y})")
        return True
    except Exception as e:
        print(f"첫 행 선택 메시지 전송 실패: {e}")
        return False


def click_confirm_hwnd_fast(confirm_hwnd):
    try:
        win32gui.SendMessage(confirm_hwnd, win32con.BM_CLICK, 0, 0)
        return True
    except Exception as e:
        print(f"확인 버튼 BM_CLICK 실패: {e}")
        return False


def wait_frame_closed(frame_hwnd, timeout=0.65):
    end = time.time() + timeout
    while time.time() < end:
        if not is_hwnd_alive_visible(frame_hwnd):
            return True
        time.sleep(0.045)
    return not is_hwnd_alive_visible(frame_hwnd)


def fallback_click_first_group_and_confirm(frame_hwnd):
    """v48 빠른 선택창 처리: Win32 핸들로 트리/확인 버튼을 바로 처리."""
    tree_hwnd = find_group_tree_hwnd_fast(frame_hwnd)
    if not tree_hwnd:
        print("조직도 트리 hwnd를 찾지 못했습니다.")
        return False

    if not send_top_row_select_message_to_tree(tree_hwnd, frame_hwnd=frame_hwnd):
        return False

    confirm_hwnd = find_button_hwnd_by_text_fast(frame_hwnd, "확인")
    if not confirm_hwnd:
        print("확인 버튼 hwnd를 찾지 못했습니다. UIA 방식으로 보조 시도합니다.")
        frame = reacquire_frame_from_hwnd(frame_hwnd)
        if frame is None:
            return True
        confirm_btn = find_in_descendants(frame, control_type="Button", name="확인", auto_id="1003")
        if confirm_btn is None:
            return False
        return invoke_confirm_and_check(frame_hwnd, confirm_btn, "사용자그룹 선택 확인 - UIA 보조")

    print("확인 버튼 빠른 실행")
    if not click_confirm_hwnd_fast(confirm_hwnd):
        return False

    if wait_frame_closed(frame_hwnd, timeout=0.65):
        print("사용자그룹 선택창 닫힘 - 성공")
        return True

    print("확인 후에도 선택창이 남아 있습니다. 선택 실패로 판단합니다.")
    return False


def select_first_group_and_confirm(frame):
    print("\n[첫 번째 사용자그룹 선택 후 확인 - v48 FAST]")
    frame_hwnd = info_of(frame).get("handle")
    if not frame_hwnd:
        print("선택 프레임 핸들을 찾지 못했습니다.")
        return False
    return fallback_click_first_group_and_confirm(frame_hwnd)



# ----------------------------------------------------------------------
# v48 OVERRIDE: 빠른 속도 유지 + 첫 번째/두 번째 결과 확인
# ----------------------------------------------------------------------
# v47은 첫 번째 결과만 확인해서 빨랐지만, 검색 결과 없음/복수/불일치 사유를 충분히 구분하지 못했다.
# v48은 전체 결과 개수 세기는 하지 않고, 첫 번째와 두 번째 결과만 읽는다.
# - 첫 번째 결과 없음: 검색 결과 없음
# - 두 번째 결과가 첫 번째와 다름: 검색 결과 복수
# - 첫 번째 이름 불일치: 대상자 이름 불일치
# - 첫 번째 부서 불일치: 부서/기관명 불일치
# - 첫 번째 일치 + 두 번째 없음/동일: 성공

def get_first_two_results_fast(result_hwnd):
    """검색 결과 리스트에서 첫 번째와 두 번째 결과만 빠르게 읽는다."""
    rows = []

    if not focus_result_fast(result_hwnd):
        return rows

    try:
        keyboard.send_keys("{HOME}")
        time.sleep(0.045)
    except Exception:
        pass

    first = compact_text(read_result_text_fast(result_hwnd))
    print(f"첫 번째 결과 확인: {first}")
    rows.append(first)

    try:
        keyboard.send_keys("{DOWN}")
        time.sleep(0.045)
    except Exception:
        pass

    second = compact_text(read_result_text_fast(result_hwnd))
    print(f"두 번째 결과 확인: {second}")
    rows.append(second)

    try:
        keyboard.send_keys("{HOME}")
        time.sleep(0.025)
    except Exception:
        pass

    return rows


def judge_target_first_two_fast(target_name, dept_keyword, rows):
    """첫 번째/두 번째 결과만으로 실패 사유를 구분한다."""
    first = rows[0] if len(rows) >= 1 else ""
    second = rows[1] if len(rows) >= 2 else ""

    first_ok = looks_like_result_text(first)
    second_ok = looks_like_result_text(second)

    if not first_ok:
        return {
            "status": "실패",
            "reason": "검색 결과 없음",
            "first": parse_result_text(first),
            "second": parse_result_text(second),
        }

    # 결과가 1개뿐이면 DOWN 후에도 같은 텍스트가 반복해서 읽히는 경우가 있다.
    # 두 번째가 결과 형식이고, 첫 번째와 다른 경우에만 복수 결과로 본다.
    if second_ok and compact_text(second) != compact_text(first):
        return {
            "status": "실패",
            "reason": "검색 결과 복수",
            "first": parse_result_text(first),
            "second": parse_result_text(second),
        }

    parsed = parse_result_text(first)

    if not name_matches(target_name, parsed):
        return {
            "status": "실패",
            "reason": "대상자 이름 불일치",
            "first": parsed,
            "second": parse_result_text(second),
        }

    if not dept_matches(dept_keyword, parsed):
        return {
            "status": "실패",
            "reason": "부서/기관명 불일치",
            "first": parsed,
            "second": parse_result_text(second),
        }

    return {
        "status": "성공",
        "reason": "단일 결과이며 이름·부서 일치",
        "first": parsed,
        "second": parse_result_text(second),
    }


def verify_target(main_hwnd, target_name, first_keyword, dept_keyword):
    """v48 빠른 검증: 첫 번째와 두 번째 결과만 확인해 실패 사유를 구분한다."""
    root = get_root(main_hwnd)

    if not click_search_tab(root):
        return None, {"status": "실패", "reason": "검색 탭 클릭 실패"}

    time.sleep(0.16)
    root = get_root(main_hwnd)

    filter_ok = set_search_filter_all_no_mouse(root)
    print(f"필터 전체 설정 결과: {filter_ok}")

    time.sleep(0.08)
    root = get_root(main_hwnd)

    search_edit = choose_main_search_edit(root)
    if search_edit is None:
        return None, {"status": "실패", "reason": "검색 입력칸 없음"}

    focus_and_replace_text(search_edit, first_keyword, "1차 검색어")

    root = get_root(main_hwnd)
    if not click_body_search_button(root):
        return None, {"status": "실패", "reason": "검색 버튼 클릭 실패"}

    print("1차 검색 결과 대기(빠른 폴링)")
    wait_result_list_fast(main_hwnd, must_contain=target_name, timeout=0.95)

    root = get_root(main_hwnd)
    search_edit = choose_main_search_edit(root)
    if search_edit is None:
        return None, {"status": "실패", "reason": "결과 내 재검색 입력칸 없음"}

    focus_and_replace_text(search_edit, dept_keyword, "결과 내 재검색어")

    root = get_root(main_hwnd)
    if not click_result_research_button(root):
        return None, {"status": "실패", "reason": "결과 내 재검색 실패"}

    print("결과 내 재검색 결과 대기(빠른 폴링)")
    chosen = wait_result_list_fast(main_hwnd, must_contain=dept_keyword, timeout=1.05)
    if chosen is None:
        chosen = find_result_list_fast(main_hwnd)

    if chosen is None:
        return None, {"status": "실패", "reason": "검색 결과 리스트 없음"}

    result_hwnd = chosen["hwnd"]
    rows = get_first_two_results_fast(result_hwnd)
    judgment = judge_target_first_two_fast(target_name, dept_keyword, rows)

    return {
        "hwnd": result_hwnd,
        "uia": None,
        "rows": rows,
        "parsed": judgment.get("first", {}),
    }, judgment

# ----------------------------------------------------------------------
# main
# ----------------------------------------------------------------------

def main():
    print("v48-fast-two-check 사용자그룹 첫 번째 그룹 생성 + 추가 테스트")
    print(f"※ 시작 시 {AUTO_GROUP_NAME} 그룹을 먼저 생성합니다.")
    print("※ 이미 같은 이름이 있으면 '동일한 이름이 존재합니다.' 감지 후 즉시 종료합니다.")
    print("※ ADD 입력 시 실제 그룹 생성 및 실제 사용자그룹 추가가 수행됩니다.")
    print("※ 메신저/선택창 위치 변경 안 함")
    print("※ 첫 번째 그룹 선택은 키보드 재시도 없이 조직도 트리 첫 행을 바로 선택합니다.")
    print("※ 컨텍스트 메뉴는 {VK_APPS} 단일 방식만 사용")
    print()

    target_name = input("대상자 이름: ").strip()
    if not target_name:
        print("대상자 이름이 없어 종료합니다.")
        return

    first_keyword = input("1차 검색어(비우면 대상자 이름 사용): ").strip()
    if not first_keyword:
        first_keyword = target_name

    dept_keyword = input("부서/기관명(결과 내 재검색어): ").strip()
    if not dept_keyword:
        print("부서/기관명이 없어 종료합니다.")
        return

    confirm = input("실제로 그룹 생성 및 첫 번째 그룹에 추가하려면 ADD 입력: ").strip()
    actually_add = confirm == "ADD"

    now = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = os.path.join(BASE_DIR, f"messenger_usergroup_add_firstgroup_v48_fast_two_check_{now}.txt")

    with open(out_path, "w", encoding="utf-8") as f:
        with redirect_stdout(f):
            print("JBEdu Messenger+ 사용자그룹 첫 번째 그룹 생성 + 추가 테스트 v48 FAST TWO-CHECK")
            print(f"생성 시각: {datetime.now()}")
            print(f"대상자 이름: {target_name}")
            print(f"1차 검색어: {first_keyword}")
            print(f"부서/기관명: {dept_keyword}")
            print(f"실제 실행 여부: {actually_add}")
            print(f"자동 생성 그룹명: {AUTO_GROUP_NAME}")
            print()

            main_hwnd = find_messenger_hwnd()
            if not main_hwnd:
                print("JBEdu Messenger+ 창을 찾지 못했습니다.")
                return

            print("메신저 창:")
            print(hwnd_info(main_hwnd))

            restore_messenger_without_move(main_hwnd)

            if not actually_add:
                print("ADD 확인값이 없어 실제 실행하지 않습니다.")
                return

            init_status, init_reason = create_auto_group_before_run(main_hwnd)

            print("\n" + "=" * 80)
            print("[초기 사용자그룹 생성 결과]")
            print("=" * 80)
            print(f"상태: {init_status}")
            print(f"사유: {init_reason}")

            if init_status == "DUPLICATE":
                print("동일한 이름이 존재하므로 전체 실행을 종료합니다.")
                return

            if init_status != "OK":
                print("사용자그룹 자동 생성 실패로 전체 실행을 종료합니다.")
                return

            verify_result, judgment = verify_target(
                main_hwnd,
                target_name,
                first_keyword,
                dept_keyword,
            )

            print("\n" + "=" * 80)
            print("[대상자 검증 결과]")
            print("=" * 80)
            print(judgment)

            if judgment["status"] != "성공":
                print("검증 실패로 사용자그룹 추가를 진행하지 않습니다.")
                return

            added, reason = add_verified_target_to_first_group(verify_result, target_name)

            print("\n" + "=" * 80)
            print("[최종 결과]")
            print("=" * 80)
            print(f"추가 시도 결과: {added}")
            print(f"사유: {reason}")
            print()
            print(f"메신저에서 {AUTO_GROUP_NAME} 그룹에 대상자가 들어갔는지 직접 확인하세요.")

    print()
    print("v34-stop-on-popup 테스트 완료")
    print(f"결과 파일: {out_path}")





import threading
import queue
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from html.parser import HTMLParser

UI_VERSION = "v3.0"
DEFAULT_GROUP_NAME = AUTO_GROUP_NAME


class QueueWriter:
    """stdout을 파일과 UI 큐에 동시에 보낼 수 있는 간단한 writer."""
    def __init__(self, file_obj=None, ui_queue=None, prefix=""):
        self.file_obj = file_obj
        self.ui_queue = ui_queue
        self.prefix = prefix
        self._buffer = ""

    def write(self, text):
        if self.file_obj:
            self.file_obj.write(text)
            self.file_obj.flush()
        if self.ui_queue:
            self._buffer += text
            while "\n" in self._buffer:
                line, self._buffer = self._buffer.split("\n", 1)
                line = line.strip()
                if line:
                    self.ui_queue.put(("log", f"{self.prefix}{line}"))

    def flush(self):
        if self.file_obj:
            self.file_obj.flush()


def split_row(line):
    if "\t" in line:
        return [p.strip() for p in line.split("\t")]
    if "/" in line:
        return [p.strip() for p in line.split("/")]
    if "," in line:
        return [p.strip() for p in line.split(",")]
    return [line.strip()]




def normalize_clipboard_table_text(text):
    """
    엑셀뿐 아니라 한글(HWP) 표 복사 텍스트도 표 형태로 해석하기 위한 정규화.
    한글 표는 환경에 따라 셀 구분자가 탭이 아니라 \x07, \x0b 등으로 들어오는 경우가 있어
    이를 탭으로 바꿔서 처리한다.
    """
    if text is None:
        return ""
    text = str(text)
    text = text.replace("\ufeff", "").replace("\u200b", "")
    text = text.replace("\xa0", " ")
    text = text.replace("\r\n", "\n").replace("\r", "\n")

    # 한글/리치텍스트 계열 표 복사에서 자주 보이는 셀 구분 제어문자들
    for ch in ("\x07", "\x0b", "\x1d", "\x1e", "\x1f"):
        text = text.replace(ch, "\t")

    # 줄 끝 탭/공백 정리
    lines = [line.rstrip(" \t") for line in text.split("\n")]
    return "\n".join(lines).strip(" \t\n")


def clipboard_text_has_table_shape(text):
    """셀 편집 중 Ctrl+V를 눌러도 표 붙여넣기로 처리할지 판별한다."""
    if text is None:
        return False
    raw = str(text)
    if any(ch in raw for ch in ("\t", "\n", "\r", "\x07", "\x0b", "\x1d", "\x1e", "\x1f")):
        return True
    normalized = normalize_clipboard_table_text(raw)
    if "\t" in normalized or "\n" in normalized:
        return True
    # 한글 표가 공백 여러 개로 열을 구분해서 들어오는 경우
    return bool(re.search(r"\S\s{2,}\S", normalized))


def split_table_line(line):
    line = safe_text(line)
    if not line:
        return []

    if "\t" in line:
        return [p.strip() for p in line.split("\t")]

    if "/" in line:
        return [p.strip() for p in line.split("/")]

    if "," in line:
        return [p.strip() for p in line.split(",")]

    # 한글 표가 일반 텍스트로 복사되며 열 사이가 여러 공백으로 변하는 경우 대응
    parts = re.split(r"\s{2,}", line.strip())
    if len(parts) >= 2:
        return [p.strip() for p in parts]

    return [line.strip()]


def is_probable_header_row(parts):
    joined = normalize(" ".join(parts))
    header_words = [
        "대상자명", "수신자명", "성명", "이름", "부서", "기관명", "부서기관명",
        "성공여부", "실패사유", "처리결과", "개별문구", "검색어", "아이디", "첨부파일", "첨부파일경로",
    ]
    return any(word in joined for word in header_words) and not any(ch.isdigit() for ch in joined)


def parse_clipboard_table_rows(text, min_cols=2, max_cols=None):
    """
    엑셀/한글 표에서 복사한 내용을 행·열로 해석한다.
    - 정상 탭/줄바꿈 표
    - 한글 표의 제어문자 기반 셀 구분
    - 전체가 한 줄로 들어온 경우: min_cols 단위로 다시 행 분리
    """
    text = normalize_clipboard_table_text(text)
    if not text:
        return []

    parsed = []
    for raw in text.split("\n"):
        raw = raw.strip(" \t")
        if not raw:
            continue
        parts = [p for p in split_table_line(raw) if p != ""]
        if parts:
            parsed.append(parts)

    if not parsed:
        return []

    # 한글에서 전체 표가 줄바꿈 없이 한 줄로 들어온 경우: A B A B ... 형태를 행으로 복원
    if len(parsed) == 1 and len(parsed[0]) > min_cols:
        flat = parsed[0]
        chunk = max_cols or min_cols
        rebuilt = []
        for i in range(0, len(flat), chunk):
            row = flat[i:i + chunk]
            if len(row) >= min_cols:
                rebuilt.append(row)
        if rebuilt:
            parsed = rebuilt

    # 한글에서 각 셀이 한 줄처럼 들어온 경우: A\nB\nA\nB ... 형태를 행으로 복원
    if len(parsed) > 1 and all(len(row) == 1 for row in parsed):
        flat = [row[0] for row in parsed]
        chunk = max_cols or min_cols
        rebuilt = []
        for i in range(0, len(flat), chunk):
            row = flat[i:i + chunk]
            if len(row) >= min_cols:
                rebuilt.append(row)
        if rebuilt:
            parsed = rebuilt

    # 표 제목 행이 같이 복사된 경우 제거
    if parsed and is_probable_header_row(parsed[0]):
        parsed = parsed[1:]

    return parsed


def is_header_cell_token(value):
    """한글 표가 셀 단위로 풀려 들어왔을 때 앞쪽 제목 셀을 제거하기 위한 판별."""
    value_norm = normalize(value)
    header_tokens = {
        "순번", "번호", "대상자명", "수신자명", "성명", "이름", "검색어", "아이디",
        "부서", "기관명", "부서명", "부서기관명", "소속", "소속기관", "학교명",
        "첨부파일", "첨부파일경로", "파일경로", "경로", "개별문구", "문구", "내용",
        "성공여부", "실패사유", "처리결과", "비고",
    }
    return value_norm in header_tokens


def looks_like_custom_message(value):
    """붙여넣기 열 수 추정용: 짧은 부서명/성명보다 문장형 개별문구에 가까운지 판별."""
    value = safe_text(value)
    if len(value) >= 18:
        return True
    return bool(re.search(r"(합니다|드립니다|주세요|바랍니다|확인|안내|제출|참석|협조|감사|[.!?])", value))


def _parse_clipboard_raw_rows(text):
    text = normalize_clipboard_table_text(text)
    if not text:
        return [], []
    parsed = []
    for raw in text.split("\n"):
        raw = raw.strip(" \t")
        if not raw:
            continue
        parts = [p for p in split_table_line(raw) if p != ""]
        if parts:
            parsed.append(parts)
    header_cells = []
    if parsed and is_probable_header_row(parsed[0]):
        header_cells = parsed[0]
        parsed = parsed[1:]
    return parsed, header_cells


def infer_flat_note_chunk(flat, header_cells=None):
    """
    한글 표가 셀 단위로 풀려 들어온 경우 몇 개 셀을 한 행으로 묶을지 추정한다.
    기본은 2열(수신자명/부서)로 본다.
    첨부파일 경로나 개별문구가 명확할 때만 3열/4열로 본다.
    """
    header_norm = normalize(" ".join(header_cells or []))
    if "첨부" in header_norm and "개별" in header_norm:
        return 4
    if "첨부" in header_norm or "개별" in header_norm:
        return 3

    if len(flat) >= 4 and len(flat) % 4 == 0:
        third_values = flat[2::4]
        fourth_values = flat[3::4]
        if any(looks_like_attachment_path(v) for v in third_values):
            return 4
        if any(looks_like_custom_message(v) for v in fourth_values):
            return 4

    if len(flat) >= 3 and len(flat) % 3 == 0:
        third_values = flat[2::3]
        if any(looks_like_attachment_path(v) for v in third_values):
            return 3
        if any(looks_like_custom_message(v) for v in third_values):
            return 3

    return 2


def parse_note_clipboard_rows(text):
    """
    쪽지 발송 표 붙여넣기 전용 파서.
    한글 표에서 수신자명/부서 2열만 복사했을 때, 3번째·4번째 셀을
    첨부파일/개별문구 칸으로 밀어 넣지 않고 2열 행으로 복원한다.
    """
    parsed, header_cells = _parse_clipboard_raw_rows(text)
    if not parsed:
        return []

    if parsed and all(len(row) == 1 for row in parsed):
        flat = [row[0] for row in parsed]
        while flat and is_header_cell_token(flat[0]):
            header_cells.append(flat.pop(0))
        parsed = [[x] for x in flat]

    flat_mode = False
    flat = []
    if len(parsed) == 1 and len(parsed[0]) > 4:
        flat_mode = True
        flat = parsed[0]
    elif len(parsed) > 1 and all(len(row) == 1 for row in parsed):
        flat_mode = True
        flat = [row[0] for row in parsed]

    if flat_mode:
        while flat and is_header_cell_token(flat[0]):
            header_cells.append(flat.pop(0))
        chunk = infer_flat_note_chunk(flat, header_cells=header_cells)
        rebuilt = []
        for i in range(0, len(flat), chunk):
            row = flat[i:i + chunk]
            if row:
                rebuilt.append(row)
        parsed = rebuilt

    rows = []
    for parts in parsed:
        parts = [safe_text(p).strip() for p in parts]
        if not any(parts):
            continue
        if is_probable_header_row(parts):
            continue
        if len(parts) == 1:
            rows.append((parts[0], "", "", ""))
        elif len(parts) == 2:
            rows.append((parts[0], parts[1], "", ""))
        elif len(parts) == 3:
            third = parts[2]
            if looks_like_attachment_path(third):
                rows.append((parts[0], parts[1], "", third))
            else:
                rows.append((parts[0], parts[1], third, ""))
        else:
            third, fourth = parts[2], parts[3]
            # 새 표 순서: 수신자명 / 부서·기관명 / 개별문구 / 첨부파일
            # 단, 예전 순서로 붙여넣은 경우(3번째가 파일 경로, 4번째가 문구)는 자동 보정한다.
            if looks_like_attachment_path(third) and not looks_like_attachment_path(fourth):
                rows.append((parts[0], parts[1], fourth, third))
            else:
                rows.append((parts[0], parts[1], third, fourth))
    return rows


class ClipboardHTMLTableParser(HTMLParser):
    """클립보드 HTML 표를 행/열 그대로 읽기 위한 최소 파서."""
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.rows = []
        self._in_tr = False
        self._in_cell = False
        self._current_row = []
        self._current_cell = []

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        if tag == "tr":
            self._in_tr = True
            self._current_row = []
        elif tag in ("td", "th") and self._in_tr:
            self._in_cell = True
            self._current_cell = []

    def handle_data(self, data):
        if self._in_cell:
            self._current_cell.append(data)

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag in ("td", "th") and self._in_cell:
            text = "".join(self._current_cell)
            text = re.sub(r"\s+", " ", text.replace("\xa0", " ")).strip()
            self._current_row.append(text)
            self._in_cell = False
            self._current_cell = []
        elif tag == "tr" and self._in_tr:
            if self._current_row:
                self.rows.append(self._current_row)
            self._in_tr = False
            self._current_row = []


def get_clipboard_html_table_matrix():
    """
    엑셀/한글이 HTML 표 형식을 클립보드에 올린 경우 행·열을 그대로 읽는다.
    사용자가 선택한 열이 1열이면 1열만, 2열이면 2열만 들어오게 하기 위한 우선 경로다.
    """
    try:
        fmt = win32clipboard.RegisterClipboardFormat("HTML Format")
        win32clipboard.OpenClipboard()
        try:
            if not win32clipboard.IsClipboardFormatAvailable(fmt):
                return []
            data = win32clipboard.GetClipboardData(fmt)
        finally:
            win32clipboard.CloseClipboard()
    except Exception:
        try:
            win32clipboard.CloseClipboard()
        except Exception:
            pass
        return []

    if isinstance(data, bytes):
        html_text = data.decode("utf-8", errors="ignore")
    else:
        html_text = str(data)

    # CF_HTML 헤더가 있으면 fragment 영역만 우선 사용한다.
    frag = html_text
    m_start = re.search(r"StartFragment:(\d+)", html_text)
    m_end = re.search(r"EndFragment:(\d+)", html_text)
    if m_start and m_end:
        try:
            s, e = int(m_start.group(1)), int(m_end.group(1))
            frag = html_text[s:e]
        except Exception:
            pass

    if "<table" not in frag.lower() and "<tr" not in frag.lower():
        return []

    parser = ClipboardHTMLTableParser()
    try:
        parser.feed(frag)
    except Exception:
        return []

    rows = []
    for row in parser.rows:
        cleaned = [safe_text(cell).strip() for cell in row]
        # 완전 빈 행은 제외하되, 열 개수는 그대로 둔다.
        if any(cleaned):
            rows.append(cleaned)
    return rows


def parse_clipboard_exact_matrix(text):
    """
    붙여넣기 데이터를 엑셀 방식 그대로 행/열로 해석한다.
    중요: 더 이상 '2개씩 묶기', '3번째 값은 첨부파일인지 추정' 같은 임의 로직을 쓰지 않는다.
    - 탭은 열
    - 줄바꿈은 행
    - 탭/줄바꿈이 없으면 현재 셀 하나
    """
    text = normalize_clipboard_table_text(text)
    if not text:
        return []

    rows = []
    for raw_line in text.split("\n"):
        # 엑셀식으로 줄은 행, 탭은 열이다.
        if raw_line == "":
            rows.append([""])
            continue
        if "\t" in raw_line:
            row = [safe_text(cell).strip() for cell in raw_line.split("\t")]
        else:
            row = [safe_text(raw_line).strip()]
        rows.append(row)

    # 맨 끝의 완전 빈 행만 정리한다.
    while rows and len(rows[-1]) == 1 and rows[-1][0] == "":
        rows.pop()
    return rows


def get_clipboard_excel_matrix(fallback_text=""):
    """
    실제 붙여넣기 행렬을 가져온다.
    1순위: 클립보드 HTML 표 행렬
    2순위: 텍스트 탭/줄바꿈 행렬
    """
    matrix = get_clipboard_html_table_matrix()
    if matrix:
        return matrix
    return parse_clipboard_exact_matrix(fallback_text)

def normalize_result_status(status):
    status = safe_text(status)
    if status in ("성공", "검증성공"):
        return "성공"
    if status == "중단":
        return "중단"
    if status in ("처리중", "진행중"):
        return "처리중"
    if status:
        return "실패"
    return ""


def display_reason(reason):
    """사용자 화면에 표시할 실패 사유를 정리한다."""
    reason = safe_text(reason)
    if reason in ("대상자 이름 불일치", "수신자 검색어 불일치"):
        return "검색 결과 없음"
    return reason


def attachment_display_text(attachment_text):
    """첨부파일명을 쪽지 본문 맨 아래에 표시하기 위한 문자열을 만든다."""
    paths = split_attachment_paths(attachment_text)
    names = []
    for path in paths:
        name = os.path.basename(os.path.normpath(str(path).strip().strip('"')))
        if name:
            names.append(name)
    return "\n".join(names)


def render_note_message_from_template(template_text, item, idx):
    """공통 쪽지 내용은 줄바꿈을 보존하고, 개별문구만 치환한다.

    첨부파일은 쪽지 본문에 파일명을 적지 않고, 실제 파일만 첨부한다.
    """
    msg = "" if template_text is None else str(template_text)
    msg = msg.replace("\r\n", "\n").replace("\r", "\n").strip()

    custom = "" if item.get("custom") is None else str(item.get("custom"))
    msg = msg.replace("{개별문구}", custom)

    # 이전 버전에서 안내했던 치환값이 남아 있어도 실제 쪽지에 보이지 않도록 제거한다.
    # 첨부파일명도 본문에 넣지 않는다. 파일은 별도로 첨부만 한다.
    for token in ("{검색어}", "{이름}", "{부서}", "{순번}", "{첨부파일}"):
        msg = msg.replace(token, "")

    return msg


# ----------------------------------------------------------------------
# 쪽지 발송 자동화 함수
# ----------------------------------------------------------------------

def get_note_windows():
    result = []
    for win in Desktop(backend="uia").windows():
        try:
            d = info_of(win)
            if d["name"] == "쪽지 쓰기" and d["class_name"] == "@Messenger7_Wnd":
                result.append(win)
        except Exception:
            pass
    return result


def get_note_window_handles():
    handles = set()
    for w in get_note_windows():
        try:
            h = info_of(w).get("handle")
            if h:
                handles.add(h)
        except Exception:
            pass
    return handles


def find_note_window(existing_handles=None, timeout=5.0):
    if existing_handles is None:
        existing_handles = set()
    end = time.time() + timeout
    fallback = None
    while time.time() < end:
        wins = get_note_windows()
        for win in wins:
            d = info_of(win)
            h = d.get("handle")
            if h not in existing_handles:
                print("새 쪽지 쓰기 창 찾기 성공")
                print(d)
                return win
            fallback = win
        time.sleep(0.12)
    if fallback is not None:
        print("새 창은 아니지만 기존 쪽지 쓰기 창을 사용합니다.")
        return fallback
    print("쪽지 쓰기 창을 찾지 못했습니다.")
    return None


def find_note_body_edit(note_win):
    try:
        controls = note_win.descendants()
    except Exception:
        return None
    for ctrl in controls:
        try:
            d = info_of(ctrl)
            if not is_visible_rect(ctrl):
                continue
            if d["control_type"] == "Edit" and d["class_name"] == "RICHEDIT50W" and d["auto_id"] == "1022":
                return ctrl
        except Exception:
            pass
    candidates = []
    for ctrl in controls:
        try:
            d = info_of(ctrl)
            if is_visible_rect(ctrl) and d["control_type"] == "Edit" and d["class_name"] == "RICHEDIT50W":
                candidates.append(ctrl)
        except Exception:
            pass
    if candidates:
        candidates.sort(key=lambda c: rect_width(c), reverse=True)
        return candidates[0]
    return None


def paste_note_text(edit, text):
    print("\n[쪽지 본문 입력]")
    try:
        edit.set_focus()
        time.sleep(0.08)
    except Exception as e:
        print(f"본문 입력칸 포커스 실패: {e}")
        return False
    try:
        keyboard.send_keys("^a")
        time.sleep(0.04)
        keyboard.send_keys("{BACKSPACE}")
        time.sleep(0.04)
        # RICHEDIT50W에는 CRLF가 가장 안정적으로 줄바꿈으로 붙여넣어진다.
        note_text = "" if text is None else str(text)
        note_text = note_text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
        set_clipboard_text(note_text)
        keyboard.send_keys("^v")
        time.sleep(0.16)
        print("본문 입력 완료")
        return True
    except Exception as e:
        print(f"본문 입력 실패: {e}")
        return False


def find_button(root, button_name=None, auto_id=None):
    try:
        controls = root.descendants()
    except Exception:
        return None
    for ctrl in controls:
        try:
            d = info_of(ctrl)
            if not is_visible_rect(ctrl):
                continue
            if d["control_type"] != "Button":
                continue
            if button_name is not None and d["name"] != button_name:
                continue
            if auto_id is not None and d["auto_id"] != str(auto_id):
                continue
            return ctrl
        except Exception:
            pass
    return None


def collect_top_handles():
    handles = set()
    for win in Desktop(backend="uia").windows():
        try:
            h = info_of(win).get("handle")
            if h:
                handles.add(h)
        except Exception:
            pass
    return handles


def click_confirmation_if_any(before_handles, timeout=3.0):
    confirm_button_names = ["예", "확인", "OK", "전송"]
    end = time.time() + timeout
    while time.time() < end:
        for win in Desktop(backend="uia").windows():
            try:
                d = info_of(win)
                h = d.get("handle")
                if not h or h in before_handles:
                    continue
                if not is_visible_rect(win):
                    continue
                try:
                    buttons = win.descendants(control_type="Button")
                except Exception:
                    buttons = []
                for btn in buttons:
                    bd = info_of(btn)
                    if bd["name"] in confirm_button_names:
                        print(f"전송 확인창 버튼 실행: {bd['name']}")
                        invoke_control(btn, bd["name"])
                        return True
            except Exception:
                pass
        time.sleep(0.15)
    return False


def click_send_button(note_win, actually_send=True):
    send_btn = find_button(note_win, button_name="전송", auto_id="1020")
    if send_btn is None:
        send_btn = find_button(note_win, button_name="전송")
    if send_btn is None:
        print("전송 버튼을 찾지 못했습니다.")
        return False
    if not actually_send:
        print("실제 전송은 하지 않습니다.")
        return True
    before_handles = collect_top_handles()
    ok = invoke_control(send_btn, "전송")
    if not ok:
        return False
    time.sleep(0.3)
    click_confirmation_if_any(before_handles, timeout=3.0)
    return True


def note_keyword_matches(keyword, parsed):
    keyword_norm = normalize(keyword)
    raw_norm = normalize(parsed.get("raw", ""))
    name_norm = normalize(parsed.get("name", ""))
    return bool(keyword_norm and (keyword_norm in raw_norm or keyword_norm == name_norm or keyword_norm in name_norm))


def search_note_recipient(main_hwnd, keyword, dept_keyword=""):
    root = get_root(main_hwnd)
    if not click_search_tab(root):
        return None, {"status": "실패", "reason": "검색 탭 클릭 실패"}

    time.sleep(0.12)
    root = get_root(main_hwnd)
    set_search_filter_all_no_mouse(root)

    time.sleep(0.06)
    root = get_root(main_hwnd)
    search_edit = choose_main_search_edit(root)
    if search_edit is None:
        return None, {"status": "실패", "reason": "검색 입력칸 없음"}

    focus_and_replace_text(search_edit, keyword, "쪽지 수신자 검색어")
    root = get_root(main_hwnd)
    if not click_body_search_button(root):
        return None, {"status": "실패", "reason": "검색 버튼 클릭 실패"}

    wait_result_list_fast(main_hwnd, must_contain=keyword, timeout=0.95)

    if dept_keyword:
        root = get_root(main_hwnd)
        search_edit = choose_main_search_edit(root)
        if search_edit is None:
            return None, {"status": "실패", "reason": "결과 내 재검색 입력칸 없음"}
        focus_and_replace_text(search_edit, dept_keyword, "쪽지 결과 내 재검색어")
        root = get_root(main_hwnd)
        if not click_result_research_button(root):
            return None, {"status": "실패", "reason": "결과 내 재검색 실패"}
        chosen = wait_result_list_fast(main_hwnd, must_contain=dept_keyword, timeout=1.05)
    else:
        chosen = find_result_list_fast(main_hwnd)

    if chosen is None:
        chosen = find_result_list_fast(main_hwnd)
    if chosen is None:
        return None, {"status": "실패", "reason": "검색 결과 리스트 없음"}

    result_hwnd = chosen["hwnd"]
    rows = get_first_two_results_fast(result_hwnd)
    first = rows[0] if len(rows) >= 1 else ""
    second = rows[1] if len(rows) >= 2 else ""

    if not looks_like_result_text(first):
        return None, {"status": "실패", "reason": "검색 결과 없음", "first": parse_result_text(first), "second": parse_result_text(second)}

    if looks_like_result_text(second) and compact_text(second) != compact_text(first):
        return None, {"status": "실패", "reason": "검색 결과 복수", "first": parse_result_text(first), "second": parse_result_text(second)}

    parsed = parse_result_text(first)
    if dept_keyword and not dept_matches(dept_keyword, parsed):
        # 쪽지 발송에서는 부서로 결과 내 재검색 후 남은 대상이 입력 부서와 맞지 않으면
        # 사용자가 의도한 수신자를 찾지 못한 것으로 표시한다.
        return None, {"status": "실패", "reason": "검색 결과 없음", "first": parsed, "second": parse_result_text(second)}

    if not note_keyword_matches(keyword, parsed):
        return None, {"status": "실패", "reason": "수신자 검색어 불일치", "first": parsed, "second": parse_result_text(second)}

    return {"hwnd": result_hwnd, "uia": None, "rows": rows, "parsed": parsed}, {
        "status": "성공",
        "reason": "단일 결과 확인",
        "first": parsed,
        "second": parse_result_text(second),
    }


def open_note_window_for_result(verify_result):
    existing = get_note_window_handles()
    menu = open_context_menu_with_single_vk_apps(verify_result["hwnd"], verify_result.get("uia"), "")
    if menu is None:
        return None, "컨텍스트 메뉴 열기 실패"
    ok = select_menu_item_by_keyboard(menu, "쪽지전송")
    if not ok:
        close_context_menu_only()
        return None, "쪽지전송 메뉴 선택 실패"
    note_win = find_note_window(existing_handles=existing, timeout=5.0)
    if note_win is None:
        return None, "쪽지 쓰기 창 열기 실패"
    return note_win, "쪽지 쓰기 창 열림"



def looks_like_attachment_path(value):
    """붙여넣기 시 3번째 값이 개별문구인지 첨부파일 경로인지 추정한다."""
    value = safe_text(value)
    if not value:
        return False
    v = value.strip().strip('"')
    if not v:
        return False
    if re.match(r"^[A-Za-z]:[\\/]", v):
        return True
    if v.startswith("\\\\"):
        return True
    if "/" in v or "\\" in v:
        return True
    # 파일명만 붙여넣은 경우도 어느 정도 대응
    return bool(re.search(r"\.(pdf|hwp|hwpx|xlsx|xls|docx|doc|pptx|ppt|zip|txt|jpg|jpeg|png)$", v, re.I))


def split_attachment_paths(value):
    """
    첨부파일 경로 입력값을 목록으로 나눈다.
    - 여러 파일은 ; 또는 | 또는 줄바꿈으로 구분
    - 공백이 포함된 경로는 그대로 유지
    """
    if value is None:
        return []
    text = str(value).strip()
    if not text:
        return []
    parts = []
    for chunk in re.split(r"[\r\n;|]+", text):
        path = chunk.strip().strip('"')
        if path:
            parts.append(path)
    return parts


def resolve_attachment_path(path):
    path = os.path.expandvars(os.path.expanduser(str(path).strip().strip('"')))
    path = os.path.normpath(path)
    if not os.path.isabs(path):
        path = os.path.abspath(path)
    return path


def validate_attachment_paths(attachment_text):
    paths = split_attachment_paths(attachment_text)
    resolved = []
    missing = []
    for path in paths:
        full = resolve_attachment_path(path)
        if not os.path.isfile(full):
            missing.append(path)
        else:
            resolved.append(full)
    return resolved, missing


def find_file_dialog(before_handles=None, timeout=5.0):
    """파일첨부 버튼 클릭 후 열린 파일 선택 대화상자를 찾는다."""
    if before_handles is None:
        before_handles = set()
    end = time.time() + timeout
    fallback = None

    while time.time() < end:
        for win in Desktop(backend="uia").windows():
            try:
                d = info_of(win)
                h = d.get("handle")
                if not h:
                    continue
                if h in before_handles:
                    continue
                if not is_visible_rect(win):
                    continue
                title = safe_text(d.get("name")) or safe_text(win32gui.GetWindowText(h))
                class_name = safe_text(d.get("class_name")) or safe_text(win32gui.GetClassName(h))
                if class_name == "#32770" or title in ("열기", "파일 선택", "Open") or "열기" in title:
                    return win
                if class_name == "#32770":
                    fallback = win
            except Exception:
                pass

        try:
            fg = win32gui.GetForegroundWindow()
            if fg and fg not in before_handles and win32gui.IsWindowVisible(fg):
                title = safe_text(win32gui.GetWindowText(fg))
                class_name = safe_text(win32gui.GetClassName(fg))
                if class_name == "#32770" or title in ("열기", "파일 선택", "Open") or "열기" in title:
                    try:
                        return Desktop(backend="uia").window(handle=fg).wrapper_object()
                    except Exception:
                        pass
        except Exception:
            pass

        time.sleep(0.1)

    return fallback


def choose_file_dialog_edit(dialog):
    """파일 선택 창의 '파일 이름' 입력칸으로 보이는 Edit를 고른다."""
    try:
        controls = dialog.descendants()
    except Exception:
        return None

    edits = []
    for ctrl in controls:
        try:
            d = info_of(ctrl)
            if d["control_type"] == "Edit" and is_visible_rect(ctrl):
                r = ctrl.rectangle()
                edits.append((int(r.top), -int(r.right - r.left), ctrl, d))
        except Exception:
            pass

    if not edits:
        return None

    # 보통 파일명 입력칸은 창 하단에 있으므로 top이 큰 순서 우선, 그중 폭이 넓은 것
    edits.sort(reverse=True)
    return edits[0][2]


def set_file_dialog_path(dialog, file_path):
    """파일 선택 대화상자에 경로를 입력하고 Enter로 확정한다."""
    file_path = os.path.normpath(str(file_path).strip().strip('"'))
    try:
        dlg_hwnd = info_of(dialog).get("handle")
        if dlg_hwnd:
            try:
                win32gui.SetForegroundWindow(dlg_hwnd)
                time.sleep(0.08)
            except Exception:
                pass
    except Exception:
        dlg_hwnd = None

    edit = choose_file_dialog_edit(dialog)

    if edit is not None:
        try:
            edit.set_focus()
            time.sleep(0.08)
            keyboard.send_keys("^a")
            time.sleep(0.04)
            keyboard.send_keys("{BACKSPACE}")
            time.sleep(0.04)
            set_clipboard_text(file_path)
            keyboard.send_keys("^v")
            time.sleep(0.08)
            keyboard.send_keys("{ENTER}")
            time.sleep(0.45)
            if not dlg_hwnd or not is_hwnd_alive_visible(dlg_hwnd):
                return True
            # 파일명 입력 후에도 창이 남아 있는 경우 한 번 더 Enter
            keyboard.send_keys("{ENTER}")
            time.sleep(0.35)
            return not is_hwnd_alive_visible(dlg_hwnd)
        except Exception as e:
            print(f"파일 선택창 경로 입력 실패(Edit 방식): {e}")

    # fallback: 현재 포커스가 파일명 칸에 있을 가능성을 이용
    try:
        set_clipboard_text(file_path)
        keyboard.send_keys("^v")
        time.sleep(0.08)
        keyboard.send_keys("{ENTER}")
        time.sleep(0.45)
        return not dlg_hwnd or not is_hwnd_alive_visible(dlg_hwnd)
    except Exception as e:
        print(f"파일 선택창 경로 입력 실패(fallback): {e}")
        return False


def attach_files_to_note(note_win, attachment_text):
    """
    쪽지 쓰기 창에 첨부파일을 추가한다.
    첨부파일 경로가 비어 있으면 성공으로 간주한다.
    여러 파일은 ; 또는 | 또는 줄바꿈으로 구분한다.
    """
    paths, missing = validate_attachment_paths(attachment_text)
    if missing:
        return False, "첨부파일 없음: " + ", ".join(missing[:3])

    if not paths:
        return True, "첨부 없음"

    note_hwnd = info_of(note_win).get("handle")
    for idx, path in enumerate(paths, start=1):
        # 매번 현재 쪽지 창 wrapper를 다시 획득한다.
        if note_hwnd:
            try:
                note_win = Desktop(backend="uia").window(handle=note_hwnd).wrapper_object()
            except Exception:
                pass

        attach_btn = find_button(note_win, button_name="파일첨부", auto_id="1014")
        if attach_btn is None:
            attach_btn = find_button(note_win, button_name="파일첨부")

        if attach_btn is None:
            return False, "파일첨부 버튼 찾기 실패"

        before_handles = collect_top_handles()
        if not invoke_control(attach_btn, f"파일첨부 {idx}/{len(paths)}"):
            return False, "파일첨부 버튼 실행 실패"

        dialog = find_file_dialog(before_handles=before_handles, timeout=5.0)
        if dialog is None:
            return False, "파일 선택 창 찾기 실패"

        if not set_file_dialog_path(dialog, path):
            return False, "파일 선택 실패: " + os.path.basename(path)

        # 첨부 목록 반영 대기
        time.sleep(0.35)

    return True, f"첨부 {len(paths)}개 완료"


def send_note_to_recipient(main_hwnd, keyword, dept_keyword, note_text, attachment_text="", actually_send=True):
    verify_result, judgment = search_note_recipient(main_hwnd, keyword, dept_keyword)
    print("[쪽지 수신자 검증 결과]")
    print(judgment)
    if judgment.get("status") != "성공":
        return False, judgment.get("reason", "수신자 검증 실패")

    note_win, reason = open_note_window_for_result(verify_result)
    if note_win is None:
        return False, reason

    body_edit = find_note_body_edit(note_win)
    if body_edit is None:
        return False, "쪽지 본문 입력칸 찾기 실패"

    if not paste_note_text(body_edit, note_text):
        return False, "쪽지 본문 입력 실패"

    attached, attach_reason = attach_files_to_note(note_win, attachment_text)
    if not attached:
        return False, attach_reason

    if not click_send_button(note_win, actually_send=actually_send):
        return False, "전송 버튼 실행 실패"

    if actually_send:
        return True, "쪽지 전송 완료" + (f" / {attach_reason}" if attachment_text else "")
    return True, "쪽지 작성 완료" + (f" / {attach_reason}" if attachment_text else "")



# ----------------------------------------------------------------------
# v5.4 path normalization overrides
# ----------------------------------------------------------------------

def _windows_style_path(path):
    value = os.path.expandvars(os.path.expanduser(str(path).strip().strip('"')))
    value = os.path.normpath(value)
    if re.match(r"^[A-Za-z]:/", value) or "/" in value:
        value = value.replace("/", "\\")
    return value


def resolve_attachment_path(path):
    path = _windows_style_path(path)
    if not os.path.isabs(path):
        path = os.path.abspath(path)
    return path


def set_file_dialog_path(dialog, file_path):
    """파일 선택 대화상자에 경로를 입력하고 Enter로 확정한다."""
    file_path = _windows_style_path(file_path)
    try:
        dlg_hwnd = info_of(dialog).get("handle")
        if dlg_hwnd:
            try:
                win32gui.SetForegroundWindow(dlg_hwnd)
                time.sleep(0.08)
            except Exception:
                pass
    except Exception:
        dlg_hwnd = None

    edit = choose_file_dialog_edit(dialog)
    if edit is not None:
        try:
            edit.set_focus()
            time.sleep(0.08)
            keyboard.send_keys("^a")
            time.sleep(0.04)
            keyboard.send_keys("{BACKSPACE}")
            time.sleep(0.04)
            set_clipboard_text(file_path)
            keyboard.send_keys("^v")
            time.sleep(0.08)
            keyboard.send_keys("{ENTER}")
            time.sleep(0.45)
            if not dlg_hwnd or not is_hwnd_alive_visible(dlg_hwnd):
                return True
            keyboard.send_keys("{ENTER}")
            time.sleep(0.35)
            return not is_hwnd_alive_visible(dlg_hwnd)
        except Exception as e:
            print(f"파일 선택창 경로 입력 실패(Edit 방식): {e}")
    try:
        set_clipboard_text(file_path)
        keyboard.send_keys("^v")
        time.sleep(0.08)
        keyboard.send_keys("{ENTER}")
        time.sleep(0.45)
        return not dlg_hwnd or not is_hwnd_alive_visible(dlg_hwnd)
    except Exception as e:
        print(f"파일 선택창 경로 입력 실패(fallback): {e}")
        return False



# ======================================================================
# v3.0 쪽지 발송 확장: 서식 쪽지(RTF) · 엑셀 명단 · 여러 열 치환 · 발송 기록
# ======================================================================
import sys
import json
import csv
import base64
import hashlib
from html.parser import HTMLParser

APP_VERSION = "v3.0"


def get_app_dir():
    """설정·기록 파일 위치: exe면 exe 옆, 파이썬 실행이면 스크립트 옆."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR = get_app_dir()
NOTE_PRESET_FILE = os.path.join(APP_DIR, "쪽지서식.json")
NOTE_SENT_LOG_FILE = os.path.join(APP_DIR, "쪽지발송기록.json")
NOTE_ATTACH_DIR = os.path.join(APP_DIR, "쪽지첨부")

DEFAULT_STYLE = {"size": 11, "bold": False, "ul": False, "color": "#000000"}

# ----------------------------------------------------------------------
# 서식 덩어리(runs): [[글자, {size,bold,ul,color}], ...]
# ----------------------------------------------------------------------

def norm_style(st):
    st = st or {}
    color = str(st.get("color", "#000000") or "#000000").upper()
    if not re.fullmatch(r"#[0-9A-F]{6}", color):
        color = "#000000"
    try:
        size = int(round(float(st.get("size", 11))))
    except Exception:
        size = 11
    return {"size": max(6, min(48, size)), "bold": bool(st.get("bold")), "ul": bool(st.get("ul")), "color": color}


def style_key(st):
    st = norm_style(st)
    return (st["size"], st["bold"], st["ul"], st["color"])


def merge_runs(runs):
    out = []
    for text, st in runs or []:
        if not text:
            continue
        st = norm_style(st)
        if out and style_key(out[-1][1]) == style_key(st):
            out[-1][0] += text
        else:
            out.append([text, st])
    return out


def runs_to_text(runs):
    return "".join(t for t, _ in (runs or []))


def runs_to_chars(runs):
    chars = []
    for text, st in runs or []:
        st = norm_style(st)
        for ch in str(text):
            chars.append((ch, st))
    return chars


def chars_to_runs(chars):
    return merge_runs([[ch, st] for ch, st in chars])


PLACEHOLDER_RE = re.compile(r"\{\{([^{}\n]+)\}\}")     # 한글 메일머지와 같은 {{열 이름}}


def ph(name):
    """자리표시 글자: {{열 이름}}"""
    return "{{" + safe_text(name).strip() + "}}"


def template_placeholders(runs):
    return [m.group(1).strip() for m in PLACEHOLDER_RE.finditer(runs_to_text(runs))
            if not is_image_token(m.group(1))]


def fill_runs(runs, values):
    """{열이름}을 값으로 바꾼다. 값은 자리표시 첫 글자의 서식을 따른다. 없는 열은 그대로 둔다."""
    chars = runs_to_chars(runs)
    text = "".join(ch for ch, _ in chars)
    out = []
    pos = 0
    for m in PLACEHOLDER_RE.finditer(text):
        key = m.group(1).strip()
        if key not in values:
            continue
        out.extend(chars[pos:m.start()])
        st = chars[m.start()][1]
        out.extend((ch, st) for ch in safe_text(values.get(key, "")))
        pos = m.end()
    out.extend(chars[pos:])
    return chars_to_runs(out)


IMG_TOKEN_RE = re.compile(r"\{그림:([^{}\n]+)\}")
FILE_TOKEN_RE = re.compile(r"\{첨부:([^{}\n]+)\}")
ROW_FILE_TOKEN = "{개별첨부}"
BODY_TOKEN_RE = re.compile(r"\{(그림|첨부):([^{}\n]+)\}|\{개별첨부\}")


def is_image_token(key):
    """본문 자리표시 중 그림·첨부 표시(개별문구가 아닌 것)."""
    k = safe_text(key).strip()
    return k.startswith("그림:") or k.startswith("첨부:") or k == "개별첨부"


def split_body_parts(runs):
    """본문을 순서대로 나눈다: ("text", runs) / ("image", 그림이름) / ("file", 첨부이름) / ("rowfile", None).
    {그림:이름}·{첨부:이름}·{개별첨부} 자리에 그림·공통 첨부·개별 첨부가 들어간다."""
    chars = runs_to_chars(runs)
    text = "".join(ch for ch, _ in chars)
    parts, pos = [], 0
    for m in BODY_TOKEN_RE.finditer(text):
        if m.start() > pos:
            parts.append(("text", chars_to_runs(chars[pos:m.start()])))
        if m.group(1) == "그림":
            parts.append(("image", m.group(2).strip()))
        elif m.group(1) == "첨부":
            parts.append(("file", m.group(2).strip()))
        else:
            parts.append(("rowfile", None))
        pos = m.end()
    if pos < len(text):
        parts.append(("text", chars_to_runs(chars[pos:])))
    return parts


def set_clipboard_image(path):
    """그림 파일을 클립보드에 넣는다(CF_DIB). 쪽지 본문에 Ctrl+V로 붙이기 위함."""
    import io
    from PIL import Image
    with Image.open(path) as im:
        im = im.convert("RGB")
        buf = io.BytesIO()
        im.save(buf, "BMP")
    dib = buf.getvalue()[14:]           # BMP 파일 머리(14바이트) 제외
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32con.CF_DIB, dib)
    finally:
        win32clipboard.CloseClipboard()


# ----------------------------------------------------------------------
# 서식 붙여넣기: 한글·워드·웹에서 복사한 서식(HTML Format) → runs
# ----------------------------------------------------------------------

def read_clipboard_html():
    """클립보드의 HTML Format 전체 문자열. 없으면 ''."""
    try:
        fmt = win32clipboard.RegisterClipboardFormat("HTML Format")
        win32clipboard.OpenClipboard()
        try:
            if not win32clipboard.IsClipboardFormatAvailable(fmt):
                return ""
            data = win32clipboard.GetClipboardData(fmt)
        finally:
            win32clipboard.CloseClipboard()
    except Exception:
        try:
            win32clipboard.CloseClipboard()
        except Exception:
            pass
        return ""
    if isinstance(data, bytes):
        raw = data
        text = data.decode("utf-8", errors="ignore")
    else:
        text = str(data)
        raw = text.encode("utf-8", errors="ignore")
    m1, m2 = re.search(r"StartHTML:(-?\d+)", text), re.search(r"EndHTML:(-?\d+)", text)
    if m1 and m2 and int(m1.group(1)) >= 0:
        try:
            return raw[int(m1.group(1)):int(m2.group(1))].decode("utf-8", errors="ignore")
        except Exception:
            pass
    i = text.lower().find("<html")
    return text[i:] if i >= 0 else text


_CSS_COLORS = {"black": "#000000", "red": "#FF0000", "blue": "#0000FF", "green": "#008000", "navy": "#000080",
               "white": "#FFFFFF", "gray": "#808080", "grey": "#808080", "maroon": "#800000", "purple": "#800080"}


def _css_decls(text):
    out = {}
    for part in safe_text(text).split(";"):
        if ":" in part:
            k, v = part.split(":", 1)
            out[k.strip().lower()] = v.strip().strip('"').strip("'").strip()
    return out


def _css_color(v):
    v = safe_text(v).strip().lower()
    if re.fullmatch(r"#[0-9a-f]{6}", v):
        return v.upper()
    if re.fullmatch(r"#[0-9a-f]{3}", v):
        return ("#" + "".join(c * 2 for c in v[1:])).upper()
    m = re.fullmatch(r"rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+).*\)", v)
    if m:
        return "#%02X%02X%02X" % tuple(min(255, int(x)) for x in m.groups())
    return _CSS_COLORS.get(v)


def _css_size(v):
    m = re.match(r"\s*([\d.]+)\s*(pt|px)?", safe_text(v).lower())
    if not m:
        return None
    n = float(m.group(1))
    if m.group(2) == "px":
        n *= 0.75
    return max(6, min(48, int(round(n))))


def _apply_css(st, decls):
    st = dict(st)
    if "font-size" in decls:
        sz = _css_size(decls["font-size"])
        if sz:
            st["size"] = sz
    if "font-weight" in decls:
        w = decls["font-weight"].lower()
        st["bold"] = w in ("bold", "bolder") or (w.isdigit() and int(w) >= 600)
    if "text-decoration" in decls or "text-decoration-line" in decls:
        st["ul"] = "underline" in (decls.get("text-decoration", "") + decls.get("text-decoration-line", "")).lower()
    if "color" in decls:
        c = _css_color(decls["color"])
        if c:
            st["color"] = c
    return st


def _file_url_to_path(src):
    from urllib.parse import unquote
    s = safe_text(src).strip()
    if s.lower().startswith("file:"):
        s = re.sub(r"^file:/*", "", s, flags=re.I)
        s = unquote(s)
        if re.match(r"^[A-Za-z]:", s):
            return s.replace("/", "\\") if os.name == "nt" else s
        return ("\\\\" + s.replace("/", "\\")) if os.name == "nt" else "/" + s
    if re.match(r"^[A-Za-z]:[\\/]", s) or s.startswith("/"):
        return unquote(s)
    return ""


class _RichHTMLParser(HTMLParser):
    BLOCK = {"p", "div", "li", "tr", "h1", "h2", "h3", "h4", "h5", "h6", "table", "ul", "ol"}

    def __init__(self, base_style):
        super().__init__(convert_charrefs=True)
        self.base = dict(base_style)
        self.stack = []          # (tag, style)
        self.chars = []          # (ch, style) / ("\x00IMG", path)
        self.classes = {}        # 'p.hstyle0' / '.hstyle0' → decls
        self.skip = 0
        self.in_style = False
        self.style_buf = []
        self.images = []

    def cur(self):
        return self.stack[-1][1] if self.stack else self.base

    def _newline(self):
        if self.chars and self.chars[-1][0] != "\n":
            self.chars.append(("\n", self.cur()))

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        a = {k.lower(): (v or "") for k, v in attrs}
        if tag in ("head", "script", "title"):
            self.skip += 1
            return
        if tag == "style":
            self.in_style = True
            return
        if tag == "br":
            self.chars.append(("\n", self.cur()))
            return
        if tag == "img":
            path = _file_url_to_path(a.get("src", ""))
            if path:
                self.chars.append(("\x00IMG", path))
                self.images.append(path)
            return
        if tag in self.BLOCK:
            self._newline()
        st = dict(self.cur())
        for cls in a.get("class", "").split():
            for key in ("." + cls.lower(), tag + "." + cls.lower()):
                if key in self.classes:
                    st = _apply_css(st, self.classes[key])
        if tag in ("b", "strong"):
            st["bold"] = True
        if tag == "u":
            st["ul"] = True
        if tag == "font":
            c = _css_color(a.get("color", ""))
            if c:
                st["color"] = c
        st = _apply_css(st, _css_decls(a.get("style", "")))
        if tag not in ("meta", "link", "hr", "col", "input"):
            self.stack.append((tag, st))

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag in ("head", "script", "title"):
            self.skip = max(0, self.skip - 1)
            return
        if tag == "style":
            self.in_style = False
            css = "".join(self.style_buf).replace("<!--", " ").replace("-->", " ")
            for m in re.finditer(r"([^{}]+)\{([^}]*)\}", css):
                decls = _css_decls(m.group(2))
                for sel in m.group(1).split(","):
                    sel = sel.strip().lower()
                    if sel:
                        self.classes[sel] = decls
            self.style_buf = []
            return
        for i in range(len(self.stack) - 1, -1, -1):
            if self.stack[i][0] == tag:
                del self.stack[i:]
                break
        if tag in self.BLOCK:
            self.chars.append(("\n", self.cur()))

    def handle_data(self, data):
        if self.in_style:
            self.style_buf.append(data)
            return
        if self.skip:
            return
        if re.fullmatch(r"[ \t\r\n\f]*", data) and (not self.chars or self.chars[-1][0] == "\n"):
            return                  # 문단 사이 줄바꿈·들여쓰기(HTML 소스 모양)는 버림
        text = re.sub(r"[ \t\r\n\f]+", " ", data)
        if not text:
            return
        st = self.cur()
        for ch in text:
            self.chars.append((" " if False else (" " if ch == " " else ch), st))


def html_to_runs(html, base_style=None):
    """HTML → (runs, 그림 경로 목록). 그림 자리는 runs 안에 '\x00IMG\x00경로\x00' 표시로 남긴다."""
    p = _RichHTMLParser(norm_style(base_style or DEFAULT_STYLE))
    p.feed(html)
    p.close()
    # 줄 단위 정리: 공백만 있는 줄은 빈 줄로, 줄 앞 공백 한 칸(HTML 줄바꿈 흔적)은 제거
    lines, line = [], []
    for item in p.chars:
        if item[0] == "\n":
            lines.append(line)
            line = []
        else:
            line.append(item)
    lines.append(line)
    out = []
    for i, ln in enumerate(lines):
        if not any(c != " " for c, _ in ln if c != "\x00IMG"):
            ln = [x for x in ln if x[0] == "\x00IMG"]
        while ln and ln[0][0] == " " and len(ln) > 1 and ln[1][0] != " ":
            ln = ln[1:]
            break
        out.extend(ln)
        if i < len(lines) - 1:
            out.append(("\n", (ln[-1][1] if ln and ln[-1][0] != "\x00IMG" else norm_style(base_style or DEFAULT_STYLE))))
    # 앞뒤 빈 줄 정리
    while out and out[0][0] == "\n":
        out.pop(0)
    while out and out[-1][0] == "\n":
        out.pop()
    runs = []
    for ch, st in out:
        if ch == "\x00IMG":
            runs.append(["\x00IMG\x00" + st + "\x00", dict(DEFAULT_STYLE)])
        else:
            runs.append([ch, st])
    return merge_runs(runs), p.images


NOTE_TEMPLATE_FILE = os.path.join(APP_DIR, "쪽지내용.json")
DEFAULT_FIELD_COLUMNS = ["개별문구1"]


def default_note_template():
    return {"runs": [["안녕하세요.\n\n{{개별문구1}}\n\n감사합니다.", dict(DEFAULT_STYLE)]],
            "files": {}, "columns": list(DEFAULT_FIELD_COLUMNS)}


def load_note_template():
    try:
        with open(NOTE_TEMPLATE_FILE, "r", encoding="utf-8") as f:
            d = json.load(f)
        if isinstance(d, dict) and d.get("runs") is not None:
            d.setdefault("files", {})
            d.setdefault("columns", list(DEFAULT_FIELD_COLUMNS))
            if not d["columns"]:
                d["columns"] = list(DEFAULT_FIELD_COLUMNS)
            # 예전 {열 이름} → {{열 이름}}
            names = set(d["columns"])
            old = re.compile(r"(?<!\{)\{([^{}\n]+)\}(?!\})")
            d["runs"] = [[old.sub(lambda m: ph(m.group(1)) if m.group(1).strip() in names else m.group(0), t), st]
                         for t, st in d["runs"]]
            return d
    except Exception:
        pass
    return default_note_template()


def save_note_template(tpl):
    try:
        with open(NOTE_TEMPLATE_FILE, "w", encoding="utf-8") as f:
            json.dump(dict(tpl, version=APP_VERSION), f, ensure_ascii=False, indent=1)
        return True
    except Exception as e:
        print(f"쪽지 내용 저장 실패: {e}")
        return False


def _rtf_escape(text):
    out = []
    for ch in text:
        o = ord(ch)
        if ch == "\\":
            out.append("\\\\")
        elif ch == "{":
            out.append("\\{")
        elif ch == "}":
            out.append("\\}")
        elif ch == "\n":
            out.append("\\par\n")
        elif ch == "\t":
            out.append("\\tab ")
        elif ch == "\r":
            continue
        elif o < 128:
            out.append(ch)
        elif o > 0xFFFF:
            o -= 0x10000
            for u in (0xD800 + (o >> 10), 0xDC00 + (o & 0x3FF)):
                out.append("\\u%d?" % (u - 65536 if u > 32767 else u))
        else:
            out.append("\\u%d?" % (o - 65536 if o > 32767 else o))
    return "".join(out)


def runs_to_rtf(runs, font_name="Malgun Gothic"):
    runs = merge_runs(runs)
    colors = []
    for _, st in runs:
        if st["color"] not in colors:
            colors.append(st["color"])
    color_tbl = "".join(
        "\\red%d\\green%d\\blue%d;" % (int(c[1:3], 16), int(c[3:5], 16), int(c[5:7], 16)) for c in colors
    )
    parts = [
        "{\\rtf1\\ansi\\ansicpg949\\deff0\\uc1",
        "{\\fonttbl{\\f0\\fnil\\fcharset129 %s;}}" % font_name,
        "{\\colortbl ;%s}" % color_tbl,
        "\\viewkind4\\pard\\f0 ",
    ]
    for text, st in runs:
        ctrl = "\\plain\\f0\\fs%d\\cf%d" % (st["size"] * 2, colors.index(st["color"]) + 1)
        if st["bold"]:
            ctrl += "\\b"
        if st["ul"]:
            ctrl += "\\ul"
        parts.append("{" + ctrl + " " + _rtf_escape(text) + "}")
    parts.append("}")
    return "".join(parts)


# ----------------------------------------------------------------------
# 서식 쪽지 붙여넣기 (쪽지 서식 도우미 '서식째 복사'와 같은 방식: RTF + 유니코드 글자)
# ----------------------------------------------------------------------

def set_clipboard_rich(rtf_text, plain_text):
    fmt = win32clipboard.RegisterClipboardFormat("Rich Text Format")
    for attempt in range(5):
        try:
            win32clipboard.OpenClipboard()
            break
        except Exception:
            time.sleep(0.08)
    else:
        raise RuntimeError("클립보드 열기 실패")
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(fmt, rtf_text.encode("ascii", "replace"))
        win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, plain_text.replace("\n", "\r\n"))
    finally:
        win32clipboard.CloseClipboard()


def paste_note_rich(edit, rtf_text, plain_text):
    print("\n[쪽지 본문 입력 - 서식]")
    try:
        edit.set_focus()
        time.sleep(0.08)
    except Exception as e:
        print(f"본문 입력칸 포커스 실패: {e}")
        return False
    try:
        keyboard.send_keys("^a")
        time.sleep(0.04)
        keyboard.send_keys("{BACKSPACE}")
        time.sleep(0.04)
        set_clipboard_rich(rtf_text, plain_text)
        keyboard.send_keys("^v")
        time.sleep(0.25)
        print("서식 본문 입력 완료")
        return True
    except Exception as e:
        print(f"서식 본문 입력 실패: {e}")
        return False


def send_note_v3(main_hwnd, keyword, dept_keyword, runs, attachment_text="", actually_send=True):
    """기존 send_note_to_recipient와 같은 순서. 본문만 서식(RTF)으로 붙여넣는다."""
    verify_result, judgment = search_note_recipient(main_hwnd, keyword, dept_keyword)
    print("[쪽지 수신자 검증 결과]")
    print(judgment)
    if judgment.get("status") != "성공":
        return False, judgment.get("reason", "수신자 검증 실패")

    note_win, reason = open_note_window_for_result(verify_result)
    if note_win is None:
        return False, reason

    body_edit = find_note_body_edit(note_win)
    if body_edit is None:
        return False, "쪽지 본문 입력칸 찾기 실패"

    plain = runs_to_text(runs)
    if not paste_note_rich(body_edit, runs_to_rtf(runs), plain):
        # 서식 붙여넣기가 막히면 글자만이라도 넣는다
        if not paste_note_text(body_edit, plain):
            return False, "쪽지 본문 입력 실패"

    attached, attach_reason = attach_files_to_note(note_win, attachment_text)
    if not attached:
        return False, attach_reason

    if not click_send_button(note_win, actually_send=actually_send):
        return False, "전송 버튼 실행 실패"

    if actually_send:
        return True, "쪽지 전송 완료" + (f" / {attach_reason}" if attachment_text else "")
    return True, "쪽지 작성 완료" + (f" / {attach_reason}" if attachment_text else "")


# ----------------------------------------------------------------------
# 엑셀·CSV 명단 읽기
# ----------------------------------------------------------------------

HEADER_HINTS = ("메일", "성명", "이름", "수신자", "부서", "기관", "소속", "학교")


def _cell_text(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    if hasattr(v, "strftime"):
        try:
            return v.strftime("%Y-%m-%d")
        except Exception:
            pass
    return " ".join(str(v).replace("\r", " ").replace("\n", " ").split())


def _rows_from_xlsx(path, sheet_name=None):
    try:
        import openpyxl
    except ImportError:
        raise RuntimeError("엑셀 읽기 모듈(openpyxl)이 없습니다. CSV로 저장해 불러오세요.")
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.worksheets[0]
    rows = [[_cell_text(v) for v in r] for r in ws.iter_rows(values_only=True)]
    names = list(wb.sheetnames)
    wb.close()
    return rows, names, ws.title


def _rows_from_csv(path):
    for enc in ("utf-8-sig", "cp949", "euc-kr"):
        try:
            with open(path, "r", encoding=enc, newline="") as f:
                return [[_cell_text(v) for v in r] for r in csv.reader(f)], ["CSV"], "CSV"
        except UnicodeDecodeError:
            continue
    raise RuntimeError("CSV 글자 인코딩을 알 수 없습니다.")


def load_table_file(path, sheet_name=None):
    """반환: headers, rows[{rowno, cells{제목:값}}], sheet_names, sheet"""
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        raw, names, sheet = _rows_from_xlsx(path, sheet_name)
    elif ext == ".csv":
        raw, names, sheet = _rows_from_csv(path)
    else:
        raise RuntimeError("xlsx 또는 csv 파일만 불러올 수 있습니다. (xls는 xlsx로 저장)")
    h = -1
    for i, r in enumerate(raw[:20]):
        filled = [c for c in r if c]
        if len(filled) >= 2 and any(any(k in c for k in HEADER_HINTS) for c in filled):
            h = i
            break
    if h < 0:
        h = next((i for i, r in enumerate(raw) if any(r)), 0)
    head_raw = raw[h] if raw else []
    headers, seen = [], {}
    for i, c in enumerate(head_raw):
        name = c or f"{i + 1}열"
        if name in seen:
            seen[name] += 1
            name = f"{name}({seen[name]})"
        else:
            seen[name] = 1
        headers.append(name)
    rows = []
    for i in range(h + 1, len(raw)):
        r = raw[i]
        if not any(r):
            continue
        cells = {headers[j]: (r[j] if j < len(r) else "") for j in range(len(headers))}
        rows.append({"rowno": i + 1, "cells": cells})
    # 제목도 값도 없는 열은 버림
    keep = []
    for j, hd in enumerate(headers):
        has_title = j < len(head_raw) and bool(head_raw[j])
        if has_title or any(rw["cells"].get(hd) for rw in rows):
            keep.append(hd)
    for rw in rows:
        rw["cells"] = {k: rw["cells"].get(k, "") for k in keep}
    return keep, rows, names, sheet


def guess_column(headers, exact, contains):
    for e in exact:
        if e in headers:
            return e
    for c in contains:
        for hd in headers:
            if c in hd:
                return hd
    return ""


# ----------------------------------------------------------------------
# 발송 기록 (중간에 멈춰도 보낸 사람 건너뛰기)
# ----------------------------------------------------------------------

class NoteSentLog:
    def __init__(self, path=NOTE_SENT_LOG_FILE):
        self.path = path
        self.data = {}
        try:
            if os.path.exists(path):
                with open(path, "r", encoding="utf-8") as f:
                    self.data = json.load(f) or {}
        except Exception:
            self.data = {}

    @staticmethod
    def key(source, rowno, name, dept):
        return "|".join([safe_text(source), safe_text(rowno), normalize(name), normalize(dept)])

    def get(self, key):
        return self.data.get(key)

    def mark(self, key, info):
        self.data[key] = info
        self._save()

    def clear_source(self, source):
        prefix = safe_text(source) + "|"
        n = 0
        for k in list(self.data.keys()):
            if k.startswith(prefix):
                del self.data[k]
                n += 1
        self._save()
        return n

    def _save(self):
        try:
            tmp = self.path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(self.data, f, ensure_ascii=False, indent=0)
            os.replace(tmp, self.path)
        except Exception as e:
            print(f"발송 기록 저장 실패: {e}")



# ======================================================================
# v3.0 검색·안정성 엔진
#  - 검색 필터 '전체' 고정
#  - 검색어1 → (검색어2 결과 내 재검색) → 여러 명이면 검색어3 결과 내 재검색
#  - 찾은 사람의 글자에 쓴 검색어가 모두 들어 있는지 최종 확인 (이전 결과·다른 사람 방지)
#  - 메신저가 앞에서 사라지면 멈춤
#  - 쪽지: 새 쪽지창만 사용, 본문 입력 확인, 전송 후 창 닫힘 확인
#  - 사용자그룹: 작업용 그룹 이어서 추가/보관 후 새로 만들기, 끝나면 이름 바꾸기
# ======================================================================
import ctypes
try:
    import win32process
except Exception:  # pragma: no cover
    win32process = None

SEARCH_CLEAR_TOKEN = "zq없는검색어zq"
COUNT_RE = re.compile(r"검색\s*결과\s*\(\s*([\d,]+)\s*명\s*\)")
ENGINE_STATE = {"count_label": None}   # None: 모름, True: 읽힘, False: 안 읽힘


class StopRequested(Exception):
    pass


class ForegroundLost(Exception):
    pass


def normalize_result_status(status):
    s = safe_text(status)
    if s in ("성공", "검증성공", "확인됨"):
        return "성공"
    if s == "중단":
        return "중단"
    if s in ("처리중", "진행중"):
        return "처리중"
    if s == "확인필요":
        return "확인필요"
    return "실패" if s else ""


def hwnd_pid(hwnd):
    try:
        if win32process is None:
            return None
        return win32process.GetWindowThreadProcessId(hwnd)[1]
    except Exception:
        return None


class MessengerSession:
    """한 번의 실행 동안 메신저 창과 중지·앞창 감시를 묶어 둔다."""

    def __init__(self, main_hwnd, stop_event=None, guard=True):
        self.main = main_hwnd
        self.pid = hwnd_pid(main_hwnd)
        self.stop_event = stop_event
        self.guard_on = guard and self.pid is not None

    def check_stop(self):
        if self.stop_event is not None and self.stop_event.is_set():
            raise StopRequested()

    def foreground_ok(self):
        try:
            fg = win32gui.GetForegroundWindow()
        except Exception:
            return True
        if not fg:
            return False
        return hwnd_pid(fg) == self.pid

    def bring_front(self):
        for _ in range(3):
            restore_messenger_without_move(self.main)
            if self.foreground_ok():
                return True
            time.sleep(0.25)
        return self.foreground_ok()

    def guard(self, where=""):
        self.check_stop()
        if not self.guard_on:
            return
        end = time.time() + 1.5
        while time.time() < end:
            if self.foreground_ok():
                return
            time.sleep(0.1)
        raise ForegroundLost(where)


# ----------------------------------------------------------------------
# 검색
# ----------------------------------------------------------------------

def _ctrl_text(ctrl):
    try:
        h = info_of(ctrl).get("handle")
        if h:
            t = safe_text(wm_gettext(h))
            if t:
                return t
    except Exception:
        pass
    try:
        return safe_text(ctrl.window_text())
    except Exception:
        return ""


FILTER_ITEMS = ["전체", "이름", "부서", "직위", "일반전화", "휴대전화", "담당업무"]


def _filter_texts(fe):
    """필터 칸 글자를 여러 방법으로 읽는다. (방법, 글자) 목록."""
    out = []
    try:
        h = info_of(fe).get("handle")
        if h:
            out.append(("WM_GETTEXT", safe_text(wm_gettext(h))))
    except Exception:
        pass
    try:
        out.append(("window_text", safe_text(fe.window_text())))
    except Exception:
        pass
    try:
        out.append(("legacy", safe_text(fe.legacy_properties().get("Value", ""))))
    except Exception:
        pass
    try:
        out.append(("value", safe_text(fe.get_value())))
    except Exception:
        pass
    return out


def _filter_value(fe):
    """읽힌 글자 중 목록 항목과 같은 것을 돌려준다. 없으면 ''."""
    for _, t in _filter_texts(fe):
        n = normalize(t)
        if n in FILTER_ITEMS:
            return n
    return ""


def v3_ensure_filter_all(main_hwnd, sess=None):
    """검색창 앞 필터를 '전체'로 (마우스 없음).

    검색창에 포커스 → Shift+Tab(필터로 이동) → ↑ 여러 번.
    필터에 포커스가 있을 때 ↑는 맨 위 '전체'에서 멈춘다(돌아가지 않음).
    """
    for attempt in range(2):
        if sess is not None:
            sess.guard("검색 필터")
        root = get_root(main_hwnd)
        se = choose_main_search_edit(root)
        fe = choose_filter_child_edit(root)
        if se is None:
            print("검색창을 찾지 못함")
            return False
        try:
            se.set_focus()
            time.sleep(0.08)
            keyboard.send_keys("+{TAB}")
            time.sleep(0.08)
            focused = None
            try:
                if fe is not None:
                    focused = bool(fe.has_keyboard_focus())
            except Exception:
                focused = None
            if focused is False:
                print("Shift+Tab 뒤 필터에 포커스가 없음 — 다시 시도")
                continue
            keyboard.send_keys("{UP 8}")
            time.sleep(0.15)
            keyboard.send_keys("{TAB}")      # 필터 → 검색창으로 포커스 되돌림
            time.sleep(0.08)
        except Exception as e:
            print(f"검색 필터 설정 실패: {e}")
            continue
        root = get_root(main_hwnd)
        fe = choose_filter_child_edit(root)
        texts = _filter_texts(fe) if fe is not None else []
        cur = _filter_value(fe) if fe is not None else ""
        print(f"[검색 필터] Shift+Tab → ↑×8 후 값: {texts} / 포커스 확인: {focused}")
        if cur and cur != "전체":
            print(f"검색 필터가 '{cur}'로 남음 — 다시 시도")
            continue
        print("검색 필터 '전체' 설정" + ("" if cur else " (값은 글자로 확인 불가)"))
        return True
    print("검색 필터를 '전체'로 바꾸지 못함")
    return False


def set_search_filter_all_no_mouse(root):
    """v2 경로에서도 실제로 '전체'를 고르도록 교체."""
    try:
        h = root.handle
    except Exception:
        h = find_messenger_hwnd()
    return v3_ensure_filter_all(h)


COUNT_LABEL_RE = re.compile(r"^\s*검색\s*결과\s*(\(\s*([\d,]+)\s*명\s*\))?\s*$")


def _count_from_label(name):
    m = COUNT_LABEL_RE.match(safe_text(name))
    if not m:
        return None
    return int(m.group(2).replace(",", "")) if m.group(2) else 0   # '검색결과'만 있으면 0명


def v3_read_count(main_hwnd):
    """'검색결과 (N명)' 글자 읽기. Win32 → UIA 순. '검색결과'만 있으면 0. 못 읽으면 None."""
    if ENGINE_STATE.get("count_src") in (None, "win32"):
        for h in win32_child_windows(main_hwnd):
            try:
                if not is_visible_nonzero_hwnd(h):
                    continue
                t = get_hwnd_text_fast(h)
                n = _count_from_label(t)
                if n is None:
                    m = COUNT_RE.search(t)
                    n = int(m.group(1).replace(",", "")) if m else None
                if n is not None:
                    ENGINE_STATE["count_src"] = "win32"
                    ENGINE_STATE["count_label"] = True
                    return n
            except Exception:
                pass
    return None


def v3_list_state(main_hwnd):
    chosen = find_result_list_fast(main_hwnd)
    text = safe_text(chosen.get("text")) if chosen else ""
    return chosen, text


def v3_wait_results(main_hwnd, min_wait=0.5):
    """검색 버튼을 누른 뒤 결과가 그려질 시간을 준다.
    (결과 목록 창 글자는 검색해도 바뀌지 않아 '바뀜' 판단에 쓸 수 없음 → 고정 대기 후 결과 줄을 직접 읽음)"""
    time.sleep(min_wait)


def v3_read_rows(main_hwnd):
    chosen = find_result_list_fast(main_hwnd)
    if chosen is None:
        return None, []
    rows = get_first_two_results_fast(chosen["hwnd"])
    return chosen, rows


def _rows_state(rows, count=None):
    first = rows[0] if len(rows) >= 1 else ""
    second = rows[1] if len(rows) >= 2 else ""
    if count is not None:
        if count == 0:
            return "none"
        if count >= 2:
            return "many"
        return "one" if looks_like_result_text(first) else "none"
    if not looks_like_result_text(first):
        return "none"
    if looks_like_result_text(second) and compact_text(second) != compact_text(first):
        return "many"
    return "one"


def _person_summary(parsed):
    name = parsed.get("name", "")
    pos = parsed.get("position", "")
    dept = parsed.get("dept", "")
    return " ".join(x for x in (name, pos) if x) + (f" · {dept}" if dept else "")


def _edit_value(ctrl):
    """UIA 값으로 입력칸 글자를 읽는다. 못 읽으면 None."""
    for fn in (lambda: ctrl.legacy_properties().get("Value", None),
               lambda: ctrl.get_value()):
        try:
            v = fn()
            if v is not None:
                return safe_text(v)
        except Exception:
            pass
    return None


def v3_type_search(sess, text):
    """검색창에 글자를 넣고 실제로 들어갔는지 확인한다."""
    for attempt in range(3):
        root = get_root(sess.main)
        edit = choose_main_search_edit(root)
        if edit is None:
            raise RuntimeError("검색 입력칸 없음")
        try:
            edit.set_focus()
            time.sleep(0.05)
            if hasattr(edit, "has_keyboard_focus") and not edit.has_keyboard_focus():
                # 포커스가 필터 등에 남아 있으면 필터에서 Tab으로 검색창에 들어간다
                fe = choose_filter_child_edit(root)
                if fe is not None:
                    fe.set_focus()
                    time.sleep(0.05)
                    keyboard.send_keys("{TAB}")
                    time.sleep(0.06)
        except Exception:
            pass
        focus_and_replace_text(edit, text, "검색어")
        v = _edit_value(edit)
        print(f"[검색창 확인] 검색창 글자: {v!r}")
        if v is None:
            return True            # 값을 못 읽는 환경: 확인 없이 진행
        if normalize(v) == normalize(text):
            return True
        print(f"[검색창 확인] 넣은 글자 '{text}' / 검색창 글자 '{v}' → 다시 입력")
    return False


def _search_step(sess, kind, text, prev_text=None, prev_count=None):
    if not v3_type_search(sess, text):
        raise RuntimeError("검색창에 검색어가 들어가지 않음")
    root = get_root(sess.main)
    if kind == "search":
        if not click_body_search_button(root):
            raise RuntimeError("검색 버튼 실행 실패")
        v3_wait_results(sess.main, 0.5)
    else:
        if not click_result_research_button(root):
            raise RuntimeError("결과 내 재검색 버튼 실행 실패")
        v3_wait_results(sess.main, 0.6)
    chosen, rows = v3_read_rows(sess.main)
    count = v3_read_count(sess.main)
    _, txt = v3_list_state(sess.main)
    if count is not None:
        print(f"[검색결과 인원] {count}")
    state = _rows_state(rows, count)
    return chosen, rows, count, state, txt


def _mail_id(k):
    """메일주소면 @ 앞 아이디만 (메신저는 전체 주소로는 못 찾고 아이디로는 찾음)."""
    k = safe_text(k).strip()
    if "@" in k:
        head = k.split("@", 1)[0].strip()
        if head:
            return head, True
    return k, False


def v3_find_person_once(sess, keywords, clear_first=False):
    raw_ks = [safe_text(k).strip() for k in (list(keywords) + ["", "", ""])[:3]]
    conv = [_mail_id(k) for k in raw_ks]
    ks = [c[0] for c in conv]
    mail_ids = {c[0] for c in conv if c[1]}
    k1, k2, k3 = ks
    if not k1:
        return None, {"status": "실패", "reason": "검색어1 없음"}

    sess.guard("검색 전")
    root = get_root(sess.main)
    if not click_search_tab(root):
        return None, {"status": "실패", "reason": "검색 탭 열기 실패"}
    time.sleep(0.12)
    if not v3_ensure_filter_all(sess.main, sess):
        return None, {"status": "실패", "reason": "검색 필터를 '전체'로 바꾸지 못함"}

    _, prev_text = v3_list_state(sess.main)
    prev_count = v3_read_count(sess.main)
    if clear_first:
        _, _, prev_count, _, prev_text = _search_step(sess, "search", SEARCH_CLEAR_TOKEN, prev_text, prev_count)

    sess.guard("검색어1")
    chosen, rows, count, state, txt = _search_step(sess, "search", k1, prev_text, prev_count)
    used = [k1]
    if state == "none":
        return None, {"status": "실패", "reason": f"검색어1({k1})로 찾은 사람 없음"}

    if k2:
        sess.guard("검색어2")
        chosen, rows, count, state, txt = _search_step(sess, "research", k2, txt, count)
        used.append(k2)
        if state == "none":
            return None, {"status": "실패", "reason": f"검색어1+2({k1}, {k2})로 찾은 사람 없음"}

    if state == "many" and k3:
        sess.guard("검색어3")
        chosen, rows, count, state, txt = _search_step(sess, "research", k3, txt, count)
        used.append(k3)
        if state == "none":
            return None, {"status": "실패", "reason": f"검색어 {', '.join(used)}로 찾은 사람 없음"}

    if state == "many":
        n = f"{count}명" if count else "2명 이상"
        return None, {"status": "실패", "reason": f"검색어 {', '.join(used)}로도 여러 명({n})"}

    first = rows[0]
    parsed = parse_result_text(first)
    raw = normalize(first)
    missing = []
    for k in used:
        if k in mail_ids:
            # 메일 아이디: 결과에 메일이 보이면 '아이디@'로 정확히 확인, 안 보이면(오프라인 등) 1명 결과로 판단
            if "이메일" in raw or "@" in raw:
                if normalize(k + "@") not in raw:
                    missing.append(k)
            continue
        if normalize(k) not in raw:
            missing.append(k)
    if missing:
        return None, {"status": "재확인", "reason": f"찾은 사람 정보에 '{missing[0]}'이(가) 없음",
                      "person": _person_summary(parsed)}
    if chosen is None:
        return None, {"status": "재확인", "reason": "결과 목록을 다시 찾지 못함"}
    person = {"hwnd": chosen["hwnd"], "uia": None, "rows": rows, "parsed": parsed, "text": first}
    return person, {"status": "성공", "reason": "1명 확인", "person": _person_summary(parsed), "used": used}


def v3_find_person(sess, keywords):
    """한 번 실패(이전 결과 의심)하면 결과를 비운 뒤 한 번 더 찾는다."""
    try:
        person, judg = v3_find_person_once(sess, keywords)
    except (StopRequested, ForegroundLost):
        raise
    except Exception as e:
        print(f"검색 중 예외: {e}")
        person, judg = None, {"status": "재확인", "reason": f"검색 중 오류: {e.__class__.__name__}"}
    if judg.get("status") != "재확인":
        return person, judg
    print(f"[재확인] {judg.get('reason')} → 결과를 비우고 다시 찾음")
    try:
        person, judg2 = v3_find_person_once(sess, keywords, clear_first=True)
    except (StopRequested, ForegroundLost):
        raise
    except Exception as e:
        return None, {"status": "실패", "reason": f"검색 중 오류: {e.__class__.__name__}"}
    if judg2.get("status") == "재확인":
        judg2["status"] = "실패"
        judg2["reason"] = "결과 확인 실패: " + judg2.get("reason", "")
    return person, judg2


# ----------------------------------------------------------------------
# 쪽지
# ----------------------------------------------------------------------

EM_GETTEXTLENGTHEX = 0x0400 + 95


def v3_richedit_length(hwnd):
    """메신저 쪽지 본문(RichEdit) 글자 수. 못 읽으면 None."""
    try:
        from pywinauto.remote_memory_block import RemoteMemoryBlock
        from pywinauto.controls.hwndwrapper import HwndWrapper

        class GETTEXTLENGTHEX(ctypes.Structure):
            _fields_ = [("flags", ctypes.c_uint32), ("codepage", ctypes.c_uint32)]

        g = GETTEXTLENGTHEX(0x2 | 0x8, 1200)   # GTL_PRECISE | GTL_NUMCHARS, UTF-16
        mem = RemoteMemoryBlock(HwndWrapper(hwnd), size=64)
        try:
            mem.Write(g)
            n = win32gui.SendMessage(hwnd, EM_GETTEXTLENGTHEX, mem.Address(), 0)
        finally:
            try:
                mem.CleanUp()
            except Exception:
                pass
        return int(n)
    except Exception as e:
        print(f"본문 글자 수 읽기 실패: {e}")
        return None


def v3_messenger_popups(pid, exclude):
    out = []
    for h in win32_top_windows():
        try:
            if h in exclude or not is_visible_nonzero_hwnd(h):
                continue
            if hwnd_pid(h) != pid:
                continue
            out.append(h)
        except Exception:
            pass
    return out


def v3_press_confirm(hwnd, names=("예", "확인", "OK", "전송")):
    try:
        win = Desktop(backend="uia").window(handle=hwnd).wrapper_object()
        for b in win.descendants(control_type="Button"):
            if info_of(b)["name"] in names:
                return invoke_control(b, info_of(b)["name"])
    except Exception:
        pass
    return False


def v3_close_note_window(note_hwnd, pid):
    try:
        if not note_hwnd or not win32gui.IsWindow(note_hwnd):
            return
        before = set(win32_top_windows())
        win32gui.PostMessage(note_hwnd, win32con.WM_CLOSE, 0, 0)
        end = time.time() + 1.2
        while time.time() < end:
            if not win32gui.IsWindow(note_hwnd) or not win32gui.IsWindowVisible(note_hwnd):
                return
            for h in v3_messenger_popups(pid, before | {note_hwnd}):
                v3_press_confirm(h, ("예", "확인", "OK"))
            time.sleep(0.1)
    except Exception:
        pass


def v3_open_note_window(sess, person):
    existing = get_note_window_handles()
    name = person.get("parsed", {}).get("name", "")
    menu = open_context_menu_with_single_vk_apps(person["hwnd"], None, name)
    if menu is None:
        return None, None, "메뉴 열기 실패"
    ok, first = validate_menu_first_item(menu, name)
    if not ok:
        close_context_menu_only()
        return None, None, f"메뉴의 사람이 다름: {first}"
    if not select_menu_item_by_keyboard(menu, "쪽지전송"):
        close_context_menu_only()
        return None, None, "쪽지전송 메뉴 선택 실패"
    end = time.time() + 6.0
    while time.time() < end:
        for win in get_note_windows():
            h = info_of(win).get("handle")
            if h and h not in existing:
                return win, h, "쪽지창 열림"
        time.sleep(0.12)
    return None, None, "새 쪽지창이 열리지 않음"


def _body_to_end(body):
    body.set_focus()
    time.sleep(0.06)
    keyboard.send_keys("^{END}")
    time.sleep(0.04)


def v3_fill_note_body(sess, note_win, body, parts):
    """본문을 순서대로 채운다. parts: [("text", runs) | ("image", 그림 경로) | ("file", 첨부 경로)].
    매번 본문 끝으로 커서를 옮긴 뒤 붙여넣으므로 그림은 서식에서 넣은 자리에 들어간다."""
    body.set_focus()
    time.sleep(0.08)
    keyboard.send_keys("^a")
    time.sleep(0.04)
    keyboard.send_keys("{BACKSPACE}")
    time.sleep(0.05)
    nfile = 0
    for kind, val in parts:
        sess.guard("쪽지 본문 입력")
        if kind == "text":
            plain = runs_to_text(val)
            if not plain:
                continue
            _body_to_end(body)
            set_clipboard_rich(runs_to_rtf(val), plain)
            keyboard.send_keys("^v")
            time.sleep(0.2)
        elif kind == "image":
            _body_to_end(body)
            try:
                set_clipboard_image(val)
            except Exception as e:
                return False, f"그림 넣기 실패({os.path.basename(val)}): {e.__class__.__name__}"
            keyboard.send_keys("^v")
            time.sleep(0.35)
        elif kind == "file":
            _body_to_end(body)
            ok, why = attach_files_to_note(note_win, val)
            if not ok:
                return False, why
            nfile += 1
            time.sleep(0.2)
    return True, f"첨부 {nfile}개" if nfile else "첨부 없음"


def v3_send_note_parts(sess, person, parts, actually_send=True):
    note_win, note_hwnd, reason = v3_open_note_window(sess, person)
    if note_win is None:
        return "실패", reason
    try:
        try:
            win32gui.SetForegroundWindow(note_hwnd)
        except Exception:
            pass
        time.sleep(0.15)
        sess.guard("쪽지 본문 입력")
        body = find_note_body_edit(note_win)
        if body is None:
            v3_close_note_window(note_hwnd, sess.pid)
            return "실패", "쪽지 본문 칸을 찾지 못함"
        body_hwnd = info_of(body).get("handle")
        expect = sum(len(runs_to_text(v).replace("\r", "")) for k, v in parts if k == "text")
        filled, why = False, ""
        for attempt in range(2):
            ok, why = v3_fill_note_body(sess, note_win, body, parts)
            if not ok:
                v3_close_note_window(note_hwnd, sess.pid)
                return "실패", why
            n = v3_richedit_length(body_hwnd) if body_hwnd else None
            print(f"본문 글자 수: {n} / 기대 {expect} 이상")
            if n is None or n >= expect * 0.9:
                filled = True
                break
            print("본문이 덜 들어가서 처음부터 다시 채움")
        if not filled:
            v3_close_note_window(note_hwnd, sess.pid)
            return "실패", "쪽지 본문이 제대로 들어가지 않음"

        if not actually_send:
            v3_close_note_window(note_hwnd, sess.pid)
            return "성공", "작성까지 확인(보내지 않음)"

        sess.guard("전송 직전")
        send_btn = find_button(note_win, button_name="전송", auto_id="1020") or find_button(note_win, button_name="전송")
        if send_btn is None:
            v3_close_note_window(note_hwnd, sess.pid)
            return "실패", "전송 버튼을 찾지 못함"
        before = set(win32_top_windows())
        if not invoke_control(send_btn, "전송"):
            v3_close_note_window(note_hwnd, sess.pid)
            return "실패", "전송 버튼 실행 실패"
        pressed = set()
        end = time.time() + 5.0
        while time.time() < end:
            if not win32gui.IsWindow(note_hwnd) or not win32gui.IsWindowVisible(note_hwnd):
                return "성공", "쪽지 전송 완료" + (f" / {why}" if why != "첨부 없음" else "")
            for h in v3_messenger_popups(sess.pid, before | {note_hwnd} | pressed):
                if v3_press_confirm(h):
                    pressed.add(h)
            time.sleep(0.12)
        return "확인필요", "전송 버튼을 눌렀지만 쪽지창이 닫히지 않음(보낸쪽지함 확인)"
    except (StopRequested, ForegroundLost):
        raise
    except Exception as e:
        v3_close_note_window(note_hwnd, sess.pid)
        return "실패", f"쪽지 처리 중 오류: {e.__class__.__name__}"


def v3_send_note(sess, person, runs, attachment_text="", actually_send=True):
    parts = [("text", runs)] + [("file", p) for p in split_attachment_paths(attachment_text)]
    return v3_send_note_parts(sess, person, parts, actually_send)


# ----------------------------------------------------------------------
# 빠른 메뉴 찾기·닫기 (v3: 화면 전체 UIA 검색 대신 Win32로 메뉴 창만 확인)
# ----------------------------------------------------------------------

def _menu_hwnds():
    out = []

    def cb(h, _):
        try:
            if win32gui.GetClassName(h) == "#32768" and win32gui.IsWindowVisible(h):
                out.append(h)
        except Exception:
            pass
        return True
    try:
        win32gui.EnumWindows(cb, None)
    except Exception:
        pass
    return out


def find_context_menu():
    end = time.time() + 0.8
    while True:
        hs = _menu_hwnds()
        if hs:
            try:
                return Desktop(backend="uia").window(handle=hs[0]).wrapper_object()
            except Exception:
                pass
        if time.time() > end:
            return None
        time.sleep(0.03)


def close_context_menu_only():
    if not _menu_hwnds():
        return
    try:
        keyboard.send_keys("{ESC}")
        end = time.time() + 0.4
        while time.time() < end and _menu_hwnds():
            time.sleep(0.02)
    except Exception as e:
        print(f"메뉴 닫기 실패: {e}")


def open_context_menu_on_control(ctrl, label):
    print(f"[메뉴 열기] {label}")
    close_context_menu_only()
    try:
        ctrl.set_focus()
        time.sleep(0.05)
        keyboard.send_keys("{VK_APPS}")
    except Exception as e:
        print(f"메뉴 열기 실패: {e}")
        return None
    return find_context_menu()


# ----------------------------------------------------------------------
# 메뉴 항목 고르기 (v3: 첫 줄 이름을 메뉴 항목으로 잘못 고르지 않게)
# ----------------------------------------------------------------------

def select_menu_item_by_keyboard(menu, target_name):
    targets = target_name if isinstance(target_name, (list, tuple, set)) else [target_name]
    targets = [safe_text(t) for t in targets if safe_text(t)]
    rows = get_menu_items(menu)
    visible = [r for r in rows if safe_text(r["info"].get("name"))]
    names = [r["info"].get("name", "") for r in visible]
    print(f"\n[메뉴 선택] {', '.join(targets)} / 메뉴: {names}")
    pick = None
    for t in targets:                       # 1순위: 정확히 같은 이름
        for i, r in enumerate(visible):
            if normalize(r["info"].get("name", "")) == normalize(t):
                pick = (i, r)
                break
        if pick:
            break
    if pick is None:                        # 2순위: 포함 비교, 단 첫 줄(이름)은 제외
        for t in targets:
            tn = normalize(t)
            for i, r in enumerate(visible):
                if i == 0:
                    continue
                if len(tn) >= 4 and tn in normalize(r["info"].get("name", "")):
                    pick = (i, r)
                    break
            if pick:
                break
    if pick is None:
        print(f"목표 메뉴 항목 없음: {targets}")
        return False
    idx, row = pick
    label = row["info"].get("name", "")
    try:
        row["ctrl"].invoke()
        time.sleep(CFG.MENU_ACTION_WAIT)
        print(f"{label} 실행")
        return True
    except Exception as e:
        print(f"{label} invoke 실패, 키보드로: {e}")
    try:
        keyboard.send_keys("{HOME}")
        time.sleep(0.06)
        keyboard.send_keys("{DOWN}" * (idx + 1), pause=0.02)
        keyboard.send_keys("{ENTER}")
        time.sleep(CFG.MENU_ACTION_WAIT)
        return True
    except Exception as e:
        print(f"{label} 키보드 실행 실패: {e}")
        return False


# ----------------------------------------------------------------------
# 사용자그룹
# ----------------------------------------------------------------------

def _group_menu_header(menu):
    names = menu_names_from_rows(get_menu_items(menu))
    header = names[0] if names else ""
    is_group = menu_has_name(names, ["그룹 이름 변경", "선택 그룹 삭제"])
    return header, is_group


def v3_group_menu_at_current(sess, group_name):
    """지금 선택된 줄에서 메뉴를 열어 그 그룹이 맞으면 메뉴를 돌려준다(목록 다시 훑지 않음)."""
    root = get_root(sess.main)
    gl = find_usergroup_main_list(root)
    if gl is None:
        return None
    menu = open_context_menu_on_control(gl, "현재 줄 그룹 확인")
    if menu is None:
        return None
    header, is_group = _group_menu_header(menu)
    if is_group and normalize(header) == normalize(group_name):
        return menu
    close_context_menu_only()
    return None

def v3_open_group_menu_for(sess, group_name, max_steps=15):
    """사용자그룹 목록에서 메뉴 첫 줄(그룹 이름)이 같은 그룹을 찾는다.
    '!!!'로 시작하는 그룹은 보통 맨 위에서 4번째 줄(내 프로필 → 내 이름 → 기관 그룹 → !!!)이라 그 근처를 먼저 본다."""
    root = get_root(sess.main)
    if not click_usergroup_tab(root):
        return None, None
    time.sleep(0.25)
    root = get_root(sess.main)
    gl = find_usergroup_main_list(root)
    if gl is None:
        return None, None
    target = normalize(group_name)

    def to_top():
        gl.set_focus()
        time.sleep(0.05)
        keyboard.send_keys("{UP 45}", pause=0.004)
        time.sleep(0.12)

    def check():
        menu = open_context_menu_on_control(gl, "그룹 확인")
        if menu is None:
            return None, "", False
        header, is_group = _group_menu_header(menu)
        print(f"  줄 확인: {header}")
        if is_group and normalize(header) == target:
            return menu, header, is_group
        close_context_menu_only()
        return None, header, is_group

    # 1단계: 맨 위에서 3·2·4칸 아래(=4·3·5번째 줄)를 먼저 확인
    try:
        to_top()
        pos = 0
        for want in (3, 2, 4):
            sess.guard("사용자그룹 찾기")
            gl.set_focus()
            diff = want - pos
            if diff:
                keyboard.send_keys(("{DOWN %d}" % diff) if diff > 0 else ("{UP %d}" % -diff), pause=0.004)
                time.sleep(0.08)
            pos = want
            menu, header, _ = check()
            if menu is not None:
                print(f"그룹 찾음: {want + 1}번째 줄")
                return menu, gl
    except (StopRequested, ForegroundLost):
        raise
    except Exception as e:
        print(f"빠른 찾기 실패: {e}")

    # 2단계: 맨 위부터 한 줄씩 (펼쳐진 그룹은 접어서 구성원 줄을 건너뜀)
    print("빠른 찾기 실패 → 한 줄씩 찾기")
    try:
        to_top()
    except Exception:
        return None, gl
    for step in range(max_steps):
        sess.guard("사용자그룹 찾기")
        menu, header, is_group = check()
        if menu is not None:
            return menu, gl
        try:
            gl.set_focus()
            time.sleep(0.04)
            if is_group and not normalize(header).startswith(normalize("내 프로필")):
                keyboard.send_keys("{LEFT}")
                time.sleep(0.04)
            keyboard.send_keys("{DOWN}")
            time.sleep(0.08)
        except Exception:
            pass
    return None, gl


def v3_rename_group(sess, old_name, new_name, at_current=False):
    menu = v3_group_menu_at_current(sess, old_name) if at_current else None
    if menu is None:
        menu, gl = v3_open_group_menu_for(sess, old_name)
    if menu is None:
        return "ERROR", f"'{old_name}' 그룹을 찾지 못함"
    if not select_menu_item_by_keyboard(menu, "그룹 이름 변경"):
        close_context_menu_only()
        return "ERROR", "그룹 이름 변경 메뉴 선택 실패"
    time.sleep(0.3)
    before = collect_top_window_handles()
    try:
        keyboard.send_keys("^a")
        time.sleep(0.05)
        set_clipboard_text(new_name)
        keyboard.send_keys("^v")
        time.sleep(0.1)
        keyboard.send_keys("{ENTER}")
        time.sleep(0.25)
    except Exception as e:
        return "ERROR", f"새 이름 입력 실패: {e}"
    dup, _ = wait_duplicate_group_message(before, main_hwnd=sess.main, timeout=1.5)
    if dup:
        close_duplicate_message(dup)
        time.sleep(0.15)
        keyboard.send_keys("{ESC}")
        time.sleep(0.1)
        return "DUPLICATE", "같은 이름의 그룹이 있음"
    menu2 = v3_group_menu_at_current(sess, new_name)      # 선택은 그 그룹에 그대로 있음
    if menu2 is not None:
        close_context_menu_only()
        return "OK", f"'{new_name}'(으)로 이름 바꿈"
    return "UNSURE", f"'{new_name}'(으)로 바꿨지만 확인 못함"


def v3_rename_group_unique(sess, old_name, new_name):
    at_current = False
    for i in range(1, 8):
        name = new_name if i == 1 else f"{new_name}({i})"
        st, msg = v3_rename_group(sess, old_name, name, at_current=at_current)
        if st != "DUPLICATE":
            return st, msg, name
        at_current = True      # 같은 그룹 줄에 그대로 있으니 다음 시도는 바로 메뉴
    return "ERROR", "같은 이름이 계속 있어 이름을 바꾸지 못함", new_name


def v3_create_work_group(sess, name):
    """작업용 그룹 만들기. 반환: OK / DUPLICATE / ERROR"""
    global AUTO_GROUP_NAME
    AUTO_GROUP_NAME = name
    root = get_root(sess.main)
    if not click_usergroup_tab(root):
        return "ERROR", "사용자그룹 탭 열기 실패"
    time.sleep(CFG.MENU_OPEN_WAIT)
    root = get_root(sess.main)
    gl = find_usergroup_main_list(root)
    if gl is None:
        return "ERROR", "사용자그룹 목록을 찾지 못함"
    menu = open_group_create_menu_via_no_mouse(gl)
    if menu is None:
        return "ERROR", "사용자그룹 생성 메뉴 열기 실패"
    if not select_menu_item_by_keyboard(menu, ["사용자그룹 생성", "사용자그룹 추가"]):
        close_context_menu_only()
        return "ERROR", "사용자그룹 생성 메뉴 선택 실패"
    time.sleep(CFG.MENU_OPEN_WAIT)
    before = collect_top_window_handles()
    try:
        set_clipboard_text(name)
        keyboard.send_keys("^v")
        time.sleep(0.1)
        keyboard.send_keys("{ENTER}")
        time.sleep(CFG.MENU_OPEN_WAIT)
    except Exception as e:
        return "ERROR", f"그룹 이름 입력 실패: {e}"
    dup, _ = wait_duplicate_group_message(before, main_hwnd=sess.main, timeout=3.0)
    if dup:
        close_duplicate_message(dup)
        time.sleep(0.15)
        try:
            keyboard.send_keys("{ESC}")
        except Exception:
            pass
        return "DUPLICATE", "작업용 그룹이 이미 있음"
    return "OK", f"{name} 생성"




# ----------------------------------------------------------------------
# Excel-like grid widgets
# ----------------------------------------------------------------------


class ExcelLikeGrid(tk.Frame):
    """
    Canvas 기반 엑셀식 표.
    - 칸 정렬/스크롤/방향키 이동
    - 클릭 드래그 블록 선택
    - Ctrl+C / Ctrl+V / Delete
    - 더블클릭 또는 타이핑으로 셀 직접 편집
    """
    def __init__(self, parent, columns, rows=12, on_paste=None, row_height=28):
        super().__init__(parent, bg="#ffffff")
        self.columns_def = columns
        self.on_paste = on_paste
        self.row_height = int(row_height)
        self.header_height = 28
        self.row_header_width = 52
        self.data = []
        self.active_row = 0
        self.active_col = 0
        self.selected_row = 0
        self.selected_col = columns[0]["key"] if columns else None
        self.sel_anchor = (0, 0)
        self.selection = (0, 0, 0, 0)
        self.dragging = False
        self.editor = None
        self._edit_original = ""
        self._build()
        for _ in range(rows):
            self.add_row()
        self.focus_cell(0, 0)

    @property
    def rows(self):
        """기존 코드 호환용: len(self.rows) 형태를 위해 data를 노출."""
        return self.data

    def _char_width_to_px(self, width):
        return max(72, int(width) * 9 + 24)

    def _col_widths(self):
        return [self._char_width_to_px(col.get("width", 14)) for col in self.columns_def]

    def _total_width(self):
        return self.row_header_width + sum(self._col_widths())

    def _col_left(self, col_idx):
        return self.row_header_width + sum(self._col_widths()[:col_idx])

    def _col_right(self, col_idx):
        return self._col_left(col_idx) + self._col_widths()[col_idx]

    def _cell_bbox(self, row_idx, col_idx):
        x1 = self._col_left(col_idx)
        x2 = self._col_right(col_idx)
        y1 = row_idx * self.row_height
        y2 = y1 + self.row_height
        return x1, y1, x2, y2

    def _build(self):
        self.rowconfigure(1, weight=1)
        self.columnconfigure(0, weight=1)

        self.header_canvas = tk.Canvas(self, height=self.header_height, bg="#ffffff", highlightthickness=0)
        self.header_canvas.grid(row=0, column=0, sticky="ew")
        self.body_canvas = tk.Canvas(self, bg="#ffffff", highlightthickness=0, takefocus=1)
        self.body_canvas.grid(row=1, column=0, sticky="nsew")

        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.body_canvas.yview)
        self.vsb.grid(row=1, column=1, sticky="ns")
        self.hsb = ttk.Scrollbar(self, orient="horizontal", command=self._xview)
        self.hsb.grid(row=2, column=0, sticky="ew")
        self.body_canvas.configure(yscrollcommand=self.vsb.set, xscrollcommand=self._on_xscroll)

        self.header_canvas.bind("<Button-1>", self._on_header_click)
        self.body_canvas.bind("<Button-1>", self._on_button_1)
        self.body_canvas.bind("<B1-Motion>", self._on_drag)
        self.body_canvas.bind("<ButtonRelease-1>", self._on_release)
        self.body_canvas.bind("<Double-1>", self._on_double_click)
        self.body_canvas.bind("<Key>", self._on_key)
        self.body_canvas.bind("<Control-c>", self._copy_selection)
        self.body_canvas.bind("<Control-C>", self._copy_selection)
        self.body_canvas.bind("<Control-v>", self._paste_clipboard)
        self.body_canvas.bind("<Control-V>", self._paste_clipboard)
        self.body_canvas.bind("<Control-a>", self._select_all)
        self.body_canvas.bind("<Control-A>", self._select_all)
        self.body_canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.body_canvas.bind("<Shift-MouseWheel>", self._on_shift_mousewheel)
        self.bind("<Configure>", lambda e: self._redraw())

    def _xview(self, *args):
        self.body_canvas.xview(*args)
        first, _ = self.body_canvas.xview()
        self.header_canvas.xview_moveto(first)

    def _on_xscroll(self, first, last):
        self.hsb.set(first, last)
        try:
            self.header_canvas.xview_moveto(first)
        except Exception:
            pass

    def _on_mousewheel(self, event):
        self._commit_edit()
        self.body_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        return "break"

    def _on_shift_mousewheel(self, event):
        self._commit_edit()
        self.body_canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")
        return "break"

    def _ensure_rows(self, count):
        while len(self.data) < count:
            self.data.append({col["key"]: "" for col in self.columns_def})

    def _row_bg(self, row_idx):
        status = self.get_value(row_idx, "status") if any(c["key"] == "status" for c in self.columns_def) else ""
        norm = normalize_result_status(status)
        if norm == "성공":
            return "#eefaf1"
        if norm == "실패":
            return "#fff2f0"
        if norm == "중단":
            return "#fff9db"
        if norm == "처리중":
            return "#e8f2ff"
        return "#ffffff"

    def _is_selected(self, row_idx, col_idx):
        r1, c1, r2, c2 = self._normalized_selection()
        return r1 <= row_idx <= r2 and c1 <= col_idx <= c2

    def _normalized_selection(self):
        r1, c1, r2, c2 = self.selection
        if r1 > r2:
            r1, r2 = r2, r1
        if c1 > c2:
            c1, c2 = c2, c1
        r1 = max(0, min(r1, max(0, len(self.data) - 1)))
        r2 = max(0, min(r2, max(0, len(self.data) - 1)))
        c1 = max(0, min(c1, len(self.columns_def) - 1))
        c2 = max(0, min(c2, len(self.columns_def) - 1))
        return r1, c1, r2, c2

    def _fit_text(self, text, width):
        text = safe_text(text)
        max_chars = max(2, int((width - 10) / 8))
        if len(text) > max_chars:
            return text[:max_chars - 1] + "…"
        return text

    def _redraw(self):
        self._draw_header()
        self._draw_body()

    def _draw_header(self):
        c = self.header_canvas
        c.delete("all")
        total_w = self._total_width()
        c.configure(scrollregion=(0, 0, total_w, self.header_height))
        # row header
        c.create_rectangle(0, 0, self.row_header_width, self.header_height, fill="#dbe7f3", outline="#111111")
        c.create_text(self.row_header_width / 2, self.header_height / 2, text="번호", font=("Malgun Gothic", 10, "bold"), fill="#17324d")
        for idx, col in enumerate(self.columns_def):
            x1 = self._col_left(idx)
            x2 = self._col_right(idx)
            c.create_rectangle(x1, 0, x2, self.header_height, fill="#dbe7f3", outline="#111111")
            c.create_text((x1 + x2) / 2, self.header_height / 2, text=col.get("header", col["key"]), font=("Malgun Gothic", 10, "bold"), fill="#17324d")

    def _draw_body(self):
        c = self.body_canvas
        c.delete("all")
        total_w = self._total_width()
        total_h = max(1, len(self.data) * self.row_height)
        c.configure(scrollregion=(0, 0, total_w, total_h))
        widths = self._col_widths()
        for r, row in enumerate(self.data):
            y1 = r * self.row_height
            y2 = y1 + self.row_height
            row_bg = self._row_bg(r)
            # row number
            row_num_bg = "#d9ecff" if r == self.active_row else "#f4f7fb"
            c.create_rectangle(0, y1, self.row_header_width, y2, fill=row_num_bg, outline="#111111")
            c.create_text(self.row_header_width / 2, y1 + self.row_height / 2, text=str(r + 1), font=("Malgun Gothic", 10), fill="#000000")
            for col_idx, col in enumerate(self.columns_def):
                key = col["key"]
                x1 = self._col_left(col_idx)
                x2 = x1 + widths[col_idx]
                fill = "#d9ecff" if self._is_selected(r, col_idx) else row_bg
                c.create_rectangle(x1, y1, x2, y2, fill=fill, outline="#111111")
                value = self._fit_text(row.get(key, ""), widths[col_idx])
                anchor = "center" if key in ("status",) else "w"
                x_text = (x1 + x2) / 2 if anchor == "center" else x1 + 6
                c.create_text(x_text, y1 + self.row_height / 2, text=value, font=("Malgun Gothic", 10), fill="#000000", anchor=anchor)
        # active cell outline after cells
        if self.data and 0 <= self.active_row < len(self.data) and 0 <= self.active_col < len(self.columns_def):
            x1, y1, x2, y2 = self._cell_bbox(self.active_row, self.active_col)
            c.create_rectangle(x1 + 1, y1 + 1, x2 - 1, y2 - 1, outline="#1a73e8", width=2)
        # block outline
        if self.data:
            r1, col1, r2, col2 = self._normalized_selection()
            x1, y1, _, _ = self._cell_bbox(r1, col1)
            _, _, x2, y2 = self._cell_bbox(r2, col2)
            c.create_rectangle(x1 + 1, y1 + 1, x2 - 1, y2 - 1, outline="#2b6cb0", width=2)

    def _cell_at(self, x, y):
        cx = self.body_canvas.canvasx(x)
        cy = self.body_canvas.canvasy(y)
        if cy < 0:
            return None
        row = int(cy // self.row_height)
        if row < 0:
            return None
        if row >= len(self.data):
            row = len(self.data) - 1 if self.data else 0
        if cx < self.row_header_width:
            return row, 0
        pos = self.row_header_width
        for idx, w in enumerate(self._col_widths()):
            if pos <= cx < pos + w:
                return row, idx
            pos += w
        return row, len(self.columns_def) - 1

    def _on_header_click(self, event):
        # column selection
        self._commit_edit()
        self.body_canvas.focus_set()
        x = self.header_canvas.canvasx(event.x)
        if x < self.row_header_width:
            return "break"
        pos = self.row_header_width
        for idx, w in enumerate(self._col_widths()):
            if pos <= x < pos + w:
                self.active_col = idx
                self.active_row = min(self.active_row, max(0, len(self.data) - 1))
                self.selection = (0, idx, max(0, len(self.data) - 1), idx)
                self.selected_row = self.active_row
                self.selected_col = self.columns_def[idx]["key"]
                self._redraw()
                return "break"
            pos += w
        return "break"

    def _on_button_1(self, event):
        self._commit_edit()
        self.body_canvas.focus_set()
        cell = self._cell_at(event.x, event.y)
        if not cell:
            return "break"
        r, c = cell
        self.active_row, self.active_col = r, c
        self.selected_row = r
        self.selected_col = self.columns_def[c]["key"]
        self.sel_anchor = (r, c)
        self.selection = (r, c, r, c)
        self.dragging = True
        self._redraw()
        return "break"

    def _on_drag(self, event):
        if not self.dragging:
            return "break"
        cell = self._cell_at(event.x, event.y)
        if not cell:
            return "break"
        r, c = cell
        ar, ac = self.sel_anchor
        self.active_row, self.active_col = r, c
        self.selected_row = r
        self.selected_col = self.columns_def[c]["key"]
        self.selection = (ar, ac, r, c)
        self._redraw()
        return "break"

    def _on_release(self, event):
        self.dragging = False
        return "break"

    def _on_double_click(self, event):
        cell = self._cell_at(event.x, event.y)
        if cell:
            r, c = cell
            self.focus_cell(r, self.columns_def[c]["key"])
            self._begin_edit()
        return "break"

    def _on_key(self, event):
        if self.editor is not None:
            return None
        key = event.keysym
        if key in ("Up", "Down", "Left", "Right", "Tab", "Return"):
            return self._move_key(key, shift=bool(event.state & 0x0001))
        if key == "F2":
            self._begin_edit()
            return "break"
        if key in ("Delete", "BackSpace"):
            self._clear_selection()
            return "break"
        if len(event.char) == 1 and event.char and event.char >= " ":
            self._begin_edit(initial=event.char)
            return "break"
        return None

    def _move_key(self, key, shift=False):
        r, c = self.active_row, self.active_col
        if key == "Up":
            r -= 1
        elif key == "Down" or key == "Return":
            r += 1
        elif key == "Left":
            c -= 1
        elif key == "Right" or (key == "Tab" and not shift):
            c += 1
        elif key == "Tab" and shift:
            c -= 1
        c = max(0, min(c, len(self.columns_def) - 1))
        if r < 0:
            r = 0
        if r >= len(self.data):
            self.add_row()
        self.focus_cell(r, self.columns_def[c]["key"])
        return "break"

    def _select_all(self, event=None):
        if not self.data:
            return "break"
        self.selection = (0, 0, len(self.data) - 1, len(self.columns_def) - 1)
        self.active_row, self.active_col = 0, 0
        self.selected_row = 0
        self.selected_col = self.columns_def[0]["key"]
        self._redraw()
        return "break"

    def _editor_paste(self, event=None):
        """편집칸 안에서 Ctrl+V를 눌러도 표 형태면 그리드 붙여넣기로 처리한다."""
        try:
            text = self.clipboard_get()
        except Exception:
            return None
        if not clipboard_text_has_table_shape(text):
            return None
        if self.editor is not None:
            row = self.editor.get("row", self.active_row)
            key = self.editor.get("key", self.columns_def[self.active_col]["key"] if self.columns_def else "")
            self._cancel_edit()
            self.focus_cell(row, key)
        else:
            row = self.active_row
            key = self.columns_def[self.active_col]["key"] if self.columns_def else ""
        if self.on_paste:
            self.on_paste(self, row, key, text)
        else:
            matrix = [line.split("\t") for line in normalize_clipboard_table_text(text).split("\n")]
            self.paste_matrix(row, key, matrix)
        return "break"

    def _begin_edit(self, initial=None):
        if not self.data:
            return
        col = self.columns_def[self.active_col]
        if not col.get("editable", True):
            return
        self._commit_edit()
        key = col["key"]
        x1, y1, x2, y2 = self._cell_bbox(self.active_row, self.active_col)
        value = self.get_value(self.active_row, key) if initial is None else initial
        self._edit_original = self.get_value(self.active_row, key)
        entry = tk.Entry(self.body_canvas, font=("Malgun Gothic", 10), relief="solid", bd=1)
        entry.insert(0, value)
        if initial is None:
            entry.select_range(0, "end")
        entry.icursor("end")
        self.editor = {"entry": entry, "row": self.active_row, "col": self.active_col, "key": key, "window": None}
        win = self.body_canvas.create_window(x1 + 1, y1 + 1, window=entry, anchor="nw", width=max(10, x2 - x1 - 2), height=max(10, y2 - y1 - 2))
        self.editor["window"] = win
        entry.focus_set()
        entry.bind("<Control-v>", self._editor_paste)
        entry.bind("<Control-V>", self._editor_paste)
        entry.bind("<Return>", lambda e: self._commit_and_move("Down"))
        entry.bind("<Tab>", lambda e: self._commit_and_move("Right"))
        entry.bind("<ISO_Left_Tab>", lambda e: self._commit_and_move("Left"))
        entry.bind("<Escape>", lambda e: self._cancel_edit())
        entry.bind("<FocusOut>", lambda e: self._commit_edit())

    def _commit_and_move(self, direction):
        self._commit_edit()
        self._move_key("Right" if direction == "Right" else "Left" if direction == "Left" else "Down")
        return "break"

    def _commit_edit(self):
        if self.editor is None:
            return
        ed = self.editor
        entry = ed.get("entry")
        try:
            value = entry.get()
            self.set_value(ed["row"], ed["key"], value, redraw=False)
        except Exception:
            pass
        try:
            self.body_canvas.delete(ed.get("window"))
            entry.destroy()
        except Exception:
            pass
        self.editor = None
        self._redraw()

    def _cancel_edit(self):
        if self.editor is None:
            return "break"
        ed = self.editor
        try:
            self.set_value(ed["row"], ed["key"], self._edit_original, redraw=False)
            self.body_canvas.delete(ed.get("window"))
            ed.get("entry").destroy()
        except Exception:
            pass
        self.editor = None
        self._redraw()
        self.body_canvas.focus_set()
        return "break"

    def _clear_selection(self):
        r1, c1, r2, c2 = self._normalized_selection()
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                if self.columns_def[c].get("editable", True):
                    self.set_value(r, self.columns_def[c]["key"], "", redraw=False)
                    self.clear_result(r, redraw=False)
        self._redraw()

    def _copy_selection(self, event=None):
        r1, c1, r2, c2 = self._normalized_selection()
        lines = []
        for r in range(r1, r2 + 1):
            values = []
            for c in range(c1, c2 + 1):
                values.append(self.get_value(r, self.columns_def[c]["key"]))
            lines.append("\t".join(values))
        text = "\n".join(lines)
        try:
            self.clipboard_clear()
            self.clipboard_append(text)
        except Exception:
            pass
        return "break"

    def _paste_clipboard(self, event=None):
        try:
            text = self.clipboard_get()
        except Exception:
            return "break"
        key = self.columns_def[self.active_col]["key"] if self.columns_def else ""
        if self.on_paste:
            self.on_paste(self, self.active_row, key, text)
        else:
            matrix = [line.split("\t") for line in normalize_clipboard_table_text(text).split("\n")]
            self.paste_matrix(self.active_row, key, matrix)
        return "break"

    def editable_keys(self):
        return [c["key"] for c in self.columns_def if c.get("editable", True)]

    def get_value(self, row_idx, key):
        try:
            return safe_text(self.data[row_idx].get(key, ""))
        except Exception:
            return ""

    def set_value(self, row_idx, key, value, redraw=True):
        self._ensure_rows(row_idx + 1)
        self.data[row_idx][key] = safe_text(value)
        if redraw:
            self._redraw()

    def add_row(self, values=None):
        row = {col["key"]: "" for col in self.columns_def}
        if values:
            for k, v in values.items():
                if k in row:
                    row[k] = safe_text(v)
        self.data.append(row)
        self._redraw()
        return len(self.data) - 1

    def clear(self, rows=12):
        self._commit_edit()
        self.data = []
        self.active_row = 0
        self.active_col = 0
        self.selected_row = 0
        self.selected_col = self.columns_def[0]["key"] if self.columns_def else None
        self.selection = (0, 0, 0, 0)
        for _ in range(rows):
            self.add_row()
        self.focus_cell(0, 0)

    def delete_selected(self):
        self._commit_edit()
        if not self.data:
            return
        r1, _, r2, _ = self._normalized_selection()
        del self.data[r1:r2 + 1]
        if not self.data:
            self.add_row()
        self.active_row = min(r1, len(self.data) - 1)
        self.selected_row = self.active_row
        self.selection = (self.active_row, self.active_col, self.active_row, self.active_col)
        self._redraw()

    def iter_nonempty(self, key_columns):
        for i, row in enumerate(self.data):
            values = {col["key"]: safe_text(row.get(col["key"], "")) for col in self.columns_def}
            if any(values.get(k, "").strip() for k in key_columns):
                yield i, values

    def set_result(self, row_idx, status="", reason=""):
        self.set_value(row_idx, "status", status, redraw=False)
        self.set_value(row_idx, "reason", display_reason(reason), redraw=False)
        self.active_row = row_idx
        self.selected_row = row_idx
        self.selected_col = self.columns_def[self.active_col]["key"] if self.columns_def else None
        self.selection = (row_idx, 0, row_idx, len(self.columns_def) - 1)
        self._redraw()
        self.scroll_to_row(row_idx)

    def clear_result(self, row_idx, redraw=True):
        if 0 <= row_idx < len(self.data):
            if "status" in self.data[row_idx]:
                self.data[row_idx]["status"] = ""
            if "reason" in self.data[row_idx]:
                self.data[row_idx]["reason"] = ""
        if redraw:
            self._redraw()

    def paste_matrix(self, start_row, start_key, matrix):
        self._commit_edit()
        if not matrix:
            return
        keys = self.editable_keys()
        if not keys:
            return
        if start_key not in keys:
            start_key = keys[0]
        start_col = keys.index(start_key)
        last_row = start_row
        last_edit_key = start_key
        for r_off, row_values in enumerate(matrix):
            row_idx = start_row + r_off
            self._ensure_rows(row_idx + 1)
            last_row = row_idx
            for c_off, value in enumerate(row_values):
                col_idx = start_col + c_off
                if col_idx >= len(keys):
                    break
                key = keys[col_idx]
                self.data[row_idx][key] = safe_text(value)
                self.clear_result(row_idx, redraw=False)
                last_edit_key = key
        c1 = self._key_to_col(start_key)
        c2 = self._key_to_col(last_edit_key)
        self.active_row = start_row
        self.active_col = c1
        self.selected_row = start_row
        self.selected_col = start_key
        self.selection = (start_row, c1, last_row, c2)
        self._redraw()
        self.scroll_to_row(start_row)
        self.body_canvas.focus_set()

    def _key_to_col(self, key):
        for i, col in enumerate(self.columns_def):
            if col["key"] == key:
                return i
        return 0

    def focus_cell(self, row_idx, key_or_index=0):
        if not self.data:
            self.add_row()
        row_idx = max(0, min(int(row_idx), len(self.data) - 1))
        if isinstance(key_or_index, int):
            col_idx = max(0, min(key_or_index, len(self.columns_def) - 1))
        else:
            col_idx = self._key_to_col(key_or_index)
        self.active_row = row_idx
        self.active_col = col_idx
        self.selected_row = row_idx
        self.selected_col = self.columns_def[col_idx]["key"]
        self.selection = (row_idx, col_idx, row_idx, col_idx)
        self.scroll_to_row(row_idx)
        self.body_canvas.focus_set()
        self._redraw()

    def scroll_to_row(self, row_idx):
        try:
            total_h = max(1, len(self.data) * self.row_height)
            y = max(0, row_idx * self.row_height - self.row_height * 2)
            self.body_canvas.yview_moveto(max(0, min(1, y / total_h)))
        except Exception:
            pass

    def rows_count(self):
        return len(self.data)


class MessengerBatchHelperUI:
    def __init__(self, root):
        self.root = root
        self.root.title(f"메신저 작업 도우미 {UI_VERSION}")
        self.root.geometry("1280x900")
        self.root.minsize(1120, 760)
        self.event_q = queue.Queue()
        self.worker = None
        self.stop_event = threading.Event()
        self.running = False
        self.active_tab = "group"
        self._build_style()
        self._build_ui()
        self.root.after(100, self._poll_events)
        self.root.bind("<F12>", lambda e: self.stop())

    def _build_style(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TFrame", background="#f4f7fb")
        style.configure("Card.TFrame", background="#ffffff", relief="flat")
        style.configure("Title.TLabel", background="#f4f7fb", foreground="#17324d", font=("Malgun Gothic", 20, "bold"))
        style.configure("Sub.TLabel", background="#f4f7fb", foreground="#52616f", font=("Malgun Gothic", 10))
        style.configure("CardTitle.TLabel", background="#ffffff", foreground="#17324d", font=("Malgun Gothic", 11, "bold"))
        style.configure("TLabel", background="#ffffff", foreground="#1f2933", font=("Malgun Gothic", 10))
        style.configure("TButton", font=("Malgun Gothic", 10), padding=6)
        style.configure("Accent.TButton", font=("Malgun Gothic", 10, "bold"), padding=6)

    def _build_ui(self):
        self.outer = ttk.Frame(self.root, padding=16)
        self.outer.pack(fill="both", expand=True)
        header = ttk.Frame(self.outer)
        header.pack(fill="x", pady=(0, 10))
        ttk.Label(header, text="메신저 작업 도우미", style="Title.TLabel", anchor="center").pack(fill="x")
        ttk.Label(header, text="JBEdu Messenger+ 사용자그룹 추가 · 쪽지 발송 · F12(어느 창에서나)·중지 버튼으로 중지", style="Sub.TLabel", anchor="center").pack(fill="x", pady=(4, 0))

        self.tabbar = tk.Frame(self.outer, bg="#f4f7fb")
        self.tabbar.pack(fill="x", pady=(0, 4))
        self.tab_buttons = {}
        self.tab_buttons["group"] = tk.Button(self.tabbar, text="사용자그룹 추가", command=lambda: self.show_tab("group"), relief="flat", bd=0, padx=28, pady=12, cursor="hand2")
        self.tab_buttons["note"] = tk.Button(self.tabbar, text="쪽지 발송", command=lambda: self.show_tab("note"), relief="flat", bd=0, padx=28, pady=12, cursor="hand2")
        self.tab_buttons["group"].pack(side="left", padx=(0, 6))
        self.tab_buttons["note"].pack(side="left")

        self.content = ttk.Frame(self.outer)
        self.content.pack(fill="both", expand=True)
        self.group_tab = ttk.Frame(self.content)
        self.note_tab = ttk.Frame(self.content)
        self._build_group_tab()
        self._build_note_tab()
        self.show_tab("group")
        footer = ttk.Label(self.outer, text="관련 문의 : 정책기획과 빅데이터팀(063-239-3176)", style="Sub.TLabel", anchor="center")
        footer.pack(fill="x", pady=(8, 0))

    def show_tab(self, name):
        self.active_tab = name
        for frame in (self.group_tab, self.note_tab):
            frame.pack_forget()
        (self.group_tab if name == "group" else self.note_tab).pack(fill="both", expand=True)
        for key, btn in self.tab_buttons.items():
            selected = key == name
            btn.configure(bg="#ffffff" if selected else "#dbe7f3", fg="#0f3d68" if selected else "#52616f", font=("Malgun Gothic", 14 if selected else 11, "bold"), padx=34 if selected else 24, pady=14 if selected else 10)

    def _card_grid(self, parent, row, column=0, **grid_opts):
        frame = ttk.Frame(parent, style="Card.TFrame", padding=12)
        frame.grid(row=row, column=column, **grid_opts)
        return frame

    def _build_group_tab(self):
        container = ttk.Frame(self.group_tab, padding=10)
        container.pack(fill="both", expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(2, weight=1)

        top = self._card_grid(container, 0, sticky="ew", pady=(0, 8))
        ttk.Label(top, text="사용자그룹 추가", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(top, text="- 표 자체가 엑셀처럼 입력칸으로 구성되어 있습니다. 원하는 칸에 바로 입력하거나 엑셀·한글 표를 Ctrl+V로 붙여넣으세요.\n- 생성되는 사용자그룹명은 '!!!사용자그룹추가'로 고정됩니다.\n- 실행하면 같은 행의 성공 여부와 실패 사유가 바로 채워집니다.", justify="left").pack(anchor="w", pady=(6, 0))

        info = self._card_grid(container, 1, sticky="ew", pady=(0, 8))
        info.columnconfigure(5, weight=1)
        ttk.Label(info, text="생성되는 사용자그룹명").grid(row=0, column=0, sticky="w")
        ttk.Label(info, text=DEFAULT_GROUP_NAME, font=("Malgun Gothic", 11, "bold"), foreground="#0f3d68").grid(row=0, column=1, sticky="w", padx=(10, 18))
        ttk.Label(info, text="※ 변경 불가").grid(row=0, column=2, sticky="w")
        self.group_status_var = tk.StringVar(value="상태: 대기")
        self.group_current_var = tk.StringVar(value="현재: -")
        self.group_summary_var = tk.StringVar(value="성공 0 / 실패 0 / 중단 0")
        ttk.Label(info, textvariable=self.group_status_var).grid(row=1, column=0, sticky="w", pady=(10, 0), columnspan=2)
        ttk.Label(info, textvariable=self.group_current_var, wraplength=520).grid(row=1, column=2, sticky="w", pady=(10, 0), columnspan=2)
        ttk.Label(info, textvariable=self.group_summary_var).grid(row=1, column=4, sticky="e", pady=(10, 0), padx=(10, 0))

        card = self._card_grid(container, 2, sticky="nsew")
        card.columnconfigure(0, weight=1)
        card.rowconfigure(1, weight=1)
        ttk.Label(card, text="대상자 입력/처리 표", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        bar = ttk.Frame(card)
        bar.grid(row=0, column=1, sticky="e")
        ttk.Button(bar, text="초기화", command=self.reset_group_table).pack(side="left", padx=3)
        ttk.Button(bar, text="행 추가", command=lambda: self.group_grid.add_row()).pack(side="left", padx=3)
        ttk.Button(bar, text="선택 행 삭제", command=self.delete_group_selected).pack(side="left", padx=3)
        ttk.Button(bar, text="실패 항목만 남기기", command=self.keep_group_failures).pack(side="left", padx=3)
        ttk.Button(bar, text="전체 실행", style="Accent.TButton", command=self.start_group).pack(side="left", padx=(16, 3))
        ttk.Button(bar, text="중지", command=self.stop).pack(side="left", padx=3)
        self.group_grid = ExcelLikeGrid(card, [
            {"key": "name", "header": "대상자명", "width": 18, "editable": True},
            {"key": "dept", "header": "부서/기관명", "width": 32, "editable": True},
            {"key": "status", "header": "성공 여부", "width": 12, "editable": False},
            {"key": "reason", "header": "실패 사유 / 처리 결과", "width": 70, "editable": False},
        ], rows=12, on_paste=self._handle_group_paste)
        self.group_grid.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(8, 0))

    def _build_note_tab(self):
        container = ttk.Frame(self.note_tab, padding=10)
        container.pack(fill="both", expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(2, weight=1)

        top_area = ttk.Frame(container)
        top_area.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        top_area.columnconfigure(0, weight=1)
        top_area.columnconfigure(1, weight=1)
        guide = self._card_grid(top_area, 0, 0, sticky="nsew", padx=(0, 8))
        ttk.Label(guide, text="쪽지 발송", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(guide, text="- 표의 각 칸에 직접 입력할 수 있습니다.\n- 수신자명/부서/개별문구/첨부파일을 엑셀·한글 표에서 복사해 Ctrl+V로 붙여넣을 수 있습니다.\n- 첨부파일 칸은 직접 경로를 입력하거나, 행 선택 후 파일 선택 버튼을 사용하세요.\n- 첨부파일은 쪽지 본문에 파일명을 적지 않고 실제 파일만 첨부합니다.", justify="left").pack(anchor="w", pady=(6, 0))
        msg_card = self._card_grid(top_area, 0, 1, sticky="nsew", padx=(8, 0))
        msg_card.columnconfigure(0, weight=1)
        ttk.Label(msg_card, text="공통 쪽지 내용", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(msg_card, text="사용 가능한 치환값: {개별문구}", foreground="#52616f").grid(row=0, column=1, sticky="e")
        self.message_template = tk.Text(msg_card, height=7, wrap="word", font=("Malgun Gothic", 10), undo=True)
        self.message_template.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        self.message_template.insert("1.0", "안녕하세요.\n\n{개별문구}\n\n감사합니다.")

        info = self._card_grid(container, 1, sticky="ew", pady=(0, 8))
        info.columnconfigure(5, weight=1)
        self.note_status_var = tk.StringVar(value="상태: 대기")
        self.note_current_var = tk.StringVar(value="현재: -")
        self.note_summary_var = tk.StringVar(value="성공 0 / 실패 0 / 중단 0")
        ttk.Label(info, textvariable=self.note_status_var).grid(row=0, column=0, sticky="w", columnspan=2)
        ttk.Label(info, textvariable=self.note_current_var, wraplength=520).grid(row=0, column=2, sticky="w", columnspan=2)
        ttk.Label(info, textvariable=self.note_summary_var).grid(row=0, column=4, sticky="e", padx=(10, 0))

        card = self._card_grid(container, 2, sticky="nsew")
        card.columnconfigure(0, weight=1)
        card.rowconfigure(1, weight=1)
        ttk.Label(card, text="수신자 입력/처리 표", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        bar = ttk.Frame(card)
        bar.grid(row=0, column=1, sticky="e")
        ttk.Button(bar, text="초기화", command=self.reset_note_table).pack(side="left", padx=3)
        ttk.Button(bar, text="행 추가", command=lambda: self.note_grid.add_row()).pack(side="left", padx=3)
        ttk.Button(bar, text="선택 행 삭제", command=self.delete_note_selected).pack(side="left", padx=3)
        ttk.Button(bar, text="선택 행 첨부파일 선택", command=self.select_attachment_for_note_row).pack(side="left", padx=3)
        ttk.Button(bar, text="첫 행 미리보기", command=self.preview_note).pack(side="left", padx=3)
        ttk.Button(bar, text="실패 항목만 남기기", command=self.keep_note_failures).pack(side="left", padx=3)
        ttk.Button(bar, text="실제 발송 실행", style="Accent.TButton", command=self.start_note_send).pack(side="left", padx=(16, 3))
        ttk.Button(bar, text="중지", command=self.stop).pack(side="left", padx=3)
        self.note_grid = ExcelLikeGrid(card, [
            {"key": "keyword", "header": "수신자명", "width": 16, "editable": True},
            {"key": "dept", "header": "부서/기관명", "width": 26, "editable": True},
            {"key": "custom", "header": "개별문구", "width": 32, "editable": True},
            {"key": "attach", "header": "첨부파일", "width": 36, "editable": True},
            {"key": "status", "header": "성공 여부", "width": 12, "editable": False},
            {"key": "reason", "header": "실패 사유 / 처리 결과", "width": 48, "editable": False},
        ], rows=12, on_paste=self._handle_note_paste)
        self.note_grid.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(8, 0))

    # -----------------------------
    # Paste handlers
    # -----------------------------
    def _handle_group_paste(self, grid, start_row, start_key, text):
        if self.running:
            messagebox.showwarning("안내", "실행 중에는 붙여넣을 수 없습니다.")
            return
        # 엑셀식 그대로 붙여넣는다. 임의로 2열 단위로 묶거나 열을 추정하지 않는다.
        matrix = get_clipboard_excel_matrix(text)
        if not matrix:
            matrix = [[safe_text(text)]]
        grid.paste_matrix(start_row, start_key, matrix)
        self._update_group_summary()

    def _handle_note_paste(self, grid, start_row, start_key, text):
        if self.running:
            messagebox.showwarning("안내", "실행 중에는 붙여넣을 수 없습니다.")
            return
        # 엑셀식 그대로 붙여넣는다.
        # 수신자명 칸에 1열을 붙여넣으면 수신자명에만, 2열을 붙여넣으면 수신자명/부서에만 들어간다.
        # 첨부파일/개별문구 열 추정 로직은 사용하지 않는다.
        matrix = get_clipboard_excel_matrix(text)
        if not matrix:
            matrix = [[safe_text(text)]]
        grid.paste_matrix(start_row, start_key, matrix)
        self._update_note_summary()

    # -----------------------------
    # Group actions
    # -----------------------------
    def reset_group_table(self):
        if self.running:
            messagebox.showwarning("안내", "실행 중에는 사용할 수 없습니다.")
            return
        if messagebox.askyesno("초기화 확인", "사용자그룹 추가 표의 모든 입력값과 처리 결과를 초기화할까요?"):
            self.group_grid.clear(12)
            self.group_status_var.set("상태: 대기")
            self.group_current_var.set("현재: -")
            self.group_summary_var.set("성공 0 / 실패 0 / 중단 0")

    def delete_group_selected(self):
        if self.running:
            messagebox.showwarning("안내", "실행 중에는 사용할 수 없습니다.")
            return
        self.group_grid.delete_selected()
        if self.group_grid.rows_count() == 0:
            self.group_grid.add_row()
        self._update_group_summary()

    def _collect_group_items(self):
        items = []
        invalid = 0
        for row_idx, row in self.group_grid.iter_nonempty({"name", "dept"}):
            name = row.get("name", "").strip()
            dept = row.get("dept", "").strip()
            self.group_grid.clear_result(row_idx)
            if not name:
                self.group_grid.set_result(row_idx, "실패", "이름 누락")
                invalid += 1
                continue
            if not dept:
                self.group_grid.set_result(row_idx, "실패", "부서/기관명 누락")
                invalid += 1
                continue
            items.append({"row": row_idx, "name": name, "dept": dept})
        self._update_group_summary()
        return items, invalid

    def start_group(self):
        if self.running:
            messagebox.showwarning("안내", "이미 실행 중입니다.")
            return
        items, invalid = self._collect_group_items()
        if not items:
            messagebox.showwarning("입력 확인", "실행 가능한 대상자가 없습니다. 표의 입력값과 실패 사유를 확인하세요.")
            return
        ok = messagebox.askyesno("실행 확인", f"총 {len(items)}명을 처리합니다.\n\n시작 시 '{DEFAULT_GROUP_NAME}' 그룹을 생성하고, 성공자만 첫 번째 그룹에 추가합니다.\n계속할까요?")
        if not ok:
            return
        self.stop_event.clear()
        self.running = True
        self.group_status_var.set("상태: 실행 중")
        self.group_current_var.set("현재: 준비 중")
        self._update_group_summary()
        self.worker = threading.Thread(target=self._group_worker, args=(items,), daemon=True)
        self.worker.start()

    def keep_group_failures(self):
        if self.running:
            messagebox.showwarning("안내", "실행 중에는 사용할 수 없습니다.")
            return
        keeps = []
        for row_idx, row in self.group_grid.iter_nonempty({"name", "dept", "status", "reason"}):
            norm = normalize_result_status(row.get("status"))
            if norm and norm != "성공":
                keeps.append(row)
        self.group_grid.clear(0)
        for row in keeps:
            idx = self.group_grid.add_row({"name": row.get("name", ""), "dept": row.get("dept", ""), "status": row.get("status", ""), "reason": row.get("reason", "")})
            self.group_grid.set_result(idx, row.get("status", ""), row.get("reason", ""))
        if not keeps:
            self.group_grid.add_row()
        self._update_group_summary()
        messagebox.showinfo("안내", "실패/중단 항목만 남겼습니다.")

    # -----------------------------
    # Note actions
    # -----------------------------
    def reset_note_table(self):
        if self.running:
            messagebox.showwarning("안내", "실행 중에는 사용할 수 없습니다.")
            return
        if messagebox.askyesno("초기화 확인", "쪽지 발송 표의 모든 입력값과 처리 결과를 초기화할까요?"):
            self.note_grid.clear(12)
            self.note_status_var.set("상태: 대기")
            self.note_current_var.set("현재: -")
            self.note_summary_var.set("성공 0 / 실패 0 / 중단 0")

    def delete_note_selected(self):
        if self.running:
            messagebox.showwarning("안내", "실행 중에는 사용할 수 없습니다.")
            return
        self.note_grid.delete_selected()
        if self.note_grid.rows_count() == 0:
            self.note_grid.add_row()
        self._update_note_summary()

    def select_attachment_for_note_row(self):
        if self.running:
            messagebox.showwarning("안내", "실행 중에는 사용할 수 없습니다.")
            return
        row = self.note_grid.selected_row
        if row is None:
            messagebox.showinfo("안내", "첨부파일을 넣을 행을 먼저 선택하세요.")
            return
        files = filedialog.askopenfilenames(title="첨부파일 선택")
        if not files:
            return
        text = ";".join(_windows_style_path(p) for p in files)
        self.note_grid.set_value(row, "attach", text)
        self.note_grid.clear_result(row)
        self.note_grid.focus_cell(row, "attach")

    def _collect_note_items(self):
        items = []
        invalid = 0
        template_text = self.message_template.get("1.0", "end").strip()
        for row_idx, row in self.note_grid.iter_nonempty({"keyword", "dept", "attach", "custom"}):
            keyword = row.get("keyword", "").strip()
            dept = row.get("dept", "").strip()
            attach = row.get("attach", "").strip()
            custom = row.get("custom", "").strip()
            self.note_grid.clear_result(row_idx)
            if not keyword:
                self.note_grid.set_result(row_idx, "실패", "수신자명 누락")
                invalid += 1
                continue
            item = {"row": row_idx, "keyword": keyword, "dept": dept, "attach": attach, "custom": custom}
            item["note_text"] = render_note_message_from_template(template_text, item, row_idx + 1)
            items.append(item)
        self._update_note_summary()
        return items, invalid

    def preview_note(self):
        items, _ = self._collect_note_items()
        if not items:
            messagebox.showwarning("안내", "입력된 수신자가 없습니다.")
            return
        messagebox.showinfo("첫 번째 쪽지 미리보기", items[0].get("note_text", ""))

    def start_note_send(self):
        if self.running:
            messagebox.showwarning("안내", "이미 실행 중입니다.")
            return
        items, invalid = self._collect_note_items()
        if not items:
            messagebox.showwarning("입력 확인", "실행 가능한 수신자가 없습니다. 표의 입력값과 실패 사유를 확인하세요.")
            return
        ok = messagebox.askyesno("실제 발송 확인", f"총 {len(items)}명에게 실제로 쪽지를 발송합니다.\n\n검색 결과가 없거나 복수인 대상자는 발송하지 않습니다.\n계속할까요?")
        if not ok:
            return
        self.stop_event.clear()
        self.running = True
        self.note_status_var.set("상태: 실행 중")
        self.note_current_var.set("현재: 준비 중")
        self._update_note_summary()
        self.worker = threading.Thread(target=self._note_worker, args=(items,), daemon=True)
        self.worker.start()

    def keep_note_failures(self):
        if self.running:
            messagebox.showwarning("안내", "실행 중에는 사용할 수 없습니다.")
            return
        keeps = []
        for row_idx, row in self.note_grid.iter_nonempty({"keyword", "dept", "attach", "custom", "status", "reason"}):
            norm = normalize_result_status(row.get("status"))
            if norm and norm != "성공":
                keeps.append(row)
        self.note_grid.clear(0)
        for row in keeps:
            idx = self.note_grid.add_row({"keyword": row.get("keyword", ""), "dept": row.get("dept", ""), "attach": row.get("attach", ""), "custom": row.get("custom", ""), "status": row.get("status", ""), "reason": row.get("reason", "")})
            self.note_grid.set_result(idx, row.get("status", ""), row.get("reason", ""))
        if not keeps:
            self.note_grid.add_row()
        self._update_note_summary()
        messagebox.showinfo("안내", "실패/중단 항목만 남겼습니다.")

    # -----------------------------
    # Common / workers / events
    # -----------------------------
    def stop(self):
        if self.running:
            self.stop_event.set()
            if self.active_tab == "note":
                self.note_status_var.set("상태: 중지 요청")
                self.note_current_var.set("현재: 현재 단계 완료 후 중지")
            else:
                self.group_status_var.set("상태: 중지 요청")
                self.group_current_var.set("현재: 현재 단계 완료 후 중지")

    def _group_worker(self, items):
        global AUTO_GROUP_NAME
        AUTO_GROUP_NAME = DEFAULT_GROUP_NAME
        with open(os.devnull, "w", encoding="utf-8") as f:
            with redirect_stdout(f):
                try:
                    main_hwnd = find_messenger_hwnd()
                    if not main_hwnd:
                        self.event_q.put(("fatal", "JBEdu Messenger+ 창을 찾지 못했습니다."))
                        return
                    restore_messenger_without_move(main_hwnd)
                    self.event_q.put(("group_status", "상태: 사용자그룹 생성 중"))
                    init_status, init_reason = create_auto_group_before_run(main_hwnd)
                    if init_status == "DUPLICATE":
                        self.event_q.put(("fatal", f"동일한 이름이 존재합니다. 기존 {DEFAULT_GROUP_NAME} 그룹을 삭제한 뒤 다시 실행하세요."))
                        return
                    if init_status != "OK":
                        self.event_q.put(("fatal", f"사용자그룹 자동 생성 실패: {init_reason}"))
                        return
                    total = len(items)
                    for idx, item in enumerate(items, start=1):
                        row = item["row"]
                        if self.stop_event.is_set():
                            self.event_q.put(("group_row", {"row": row, "status": "중단", "reason": "사용자 중지"}))
                            break
                        name, dept = item["name"], item["dept"]
                        self.event_q.put(("group_current", {"row": row, "text": f"현재: {idx}/{total} - {name} / {dept}"}))
                        self.event_q.put(("group_status", "상태: 검색·검증 중"))
                        try:
                            verify_result, judgment = verify_target(main_hwnd, name, name, dept)
                            if judgment.get("status") != "성공":
                                self.event_q.put(("group_row", {"row": row, "status": "실패", "reason": judgment.get("reason", "검증 실패")}))
                                continue
                            if self.stop_event.is_set():
                                self.event_q.put(("group_row", {"row": row, "status": "중단", "reason": "사용자 중지"}))
                                break
                            self.event_q.put(("group_status", "상태: 내목록 추가 중"))
                            added, reason = add_verified_target_to_first_group(verify_result, name)
                            self.event_q.put(("group_row", {"row": row, "status": "성공" if added else "실패", "reason": reason}))
                        except Exception:
                            self.event_q.put(("group_row", {"row": row, "status": "실패", "reason": "예외 발생"}))
                    self.event_q.put(("group_done", "작업이 완료되었습니다."))
                except Exception:
                    self.event_q.put(("fatal", "실행 중 예외가 발생했습니다."))

    def _note_worker(self, items):
        with open(os.devnull, "w", encoding="utf-8") as f:
            with redirect_stdout(f):
                try:
                    main_hwnd = find_messenger_hwnd()
                    if not main_hwnd:
                        self.event_q.put(("fatal", "JBEdu Messenger+ 창을 찾지 못했습니다."))
                        return
                    restore_messenger_without_move(main_hwnd)
                    total = len(items)
                    for idx, item in enumerate(items, start=1):
                        row = item["row"]
                        if self.stop_event.is_set():
                            self.event_q.put(("note_row", {"row": row, "status": "중단", "reason": "사용자 중지"}))
                            break
                        keyword, dept = item["keyword"], item.get("dept", "")
                        self.event_q.put(("note_current", {"row": row, "text": f"현재: {idx}/{total} - {keyword} / {dept or '-'}"}))
                        self.event_q.put(("note_status", "상태: 검색·쪽지 발송 중"))
                        try:
                            sent, reason = send_note_to_recipient(main_hwnd, keyword, dept, item.get("note_text", ""), item.get("attach", ""), actually_send=True)
                            self.event_q.put(("note_row", {"row": row, "status": "성공" if sent else "실패", "reason": reason}))
                        except Exception:
                            self.event_q.put(("note_row", {"row": row, "status": "실패", "reason": "예외 발생"}))
                    self.event_q.put(("note_done", "쪽지 발송이 완료되었습니다."))
                except Exception:
                    self.event_q.put(("fatal", "실행 중 예외가 발생했습니다."))

    def _poll_events(self):
        try:
            while True:
                event, payload = self.event_q.get_nowait()
                if event == "group_status":
                    self.group_status_var.set(payload)
                elif event == "group_current":
                    self.group_current_var.set(payload["text"])
                    self.group_grid.set_result(payload["row"], "처리중", "")
                    self._update_group_summary()
                elif event == "group_row":
                    self.group_grid.set_result(payload["row"], payload.get("status", ""), payload.get("reason", ""))
                    self._update_group_summary()
                elif event == "group_done":
                    self.running = False
                    self.group_status_var.set("상태: 대기")
                    self.group_current_var.set("현재: -")
                    self._update_group_summary()
                    messagebox.showinfo("완료", self.group_summary_var.get())
                elif event == "note_status":
                    self.note_status_var.set(payload)
                elif event == "note_current":
                    self.note_current_var.set(payload["text"])
                    self.note_grid.set_result(payload["row"], "처리중", "")
                    self._update_note_summary()
                elif event == "note_row":
                    self.note_grid.set_result(payload["row"], payload.get("status", ""), payload.get("reason", ""))
                    self._update_note_summary()
                elif event == "note_done":
                    self.running = False
                    self.note_status_var.set("상태: 대기")
                    self.note_current_var.set("현재: -")
                    self._update_note_summary()
                    messagebox.showinfo("완료", self.note_summary_var.get())
                elif event == "fatal":
                    self.running = False
                    self.group_status_var.set("상태: 대기")
                    self.note_status_var.set("상태: 대기")
                    self.group_current_var.set("현재: -")
                    self.note_current_var.set("현재: -")
                    messagebox.showwarning("작업 중단", payload)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_events)

    def _update_group_summary(self):
        success = fail = stop = 0
        for _, row in self.group_grid.iter_nonempty({"name", "dept", "status", "reason"}):
            norm = normalize_result_status(row.get("status", ""))
            if norm == "성공": success += 1
            elif norm == "중단": stop += 1
            elif norm == "실패": fail += 1
        self.group_summary_var.set(f"성공 {success} / 실패 {fail} / 중단 {stop}")

    def _update_note_summary(self):
        success = fail = stop = 0
        for _, row in self.note_grid.iter_nonempty({"keyword", "dept", "attach", "custom", "status", "reason"}):
            norm = normalize_result_status(row.get("status", ""))
            if norm == "성공": success += 1
            elif norm == "중단": stop += 1
            elif norm == "실패": fail += 1
        self.note_summary_var.set(f"성공 {success} / 실패 {fail} / 중단 {stop}")


# ======================================================================
# v3.0 화면: 빠른 표 · 서식 편집기 · 새 쪽지 발송 탭
# ======================================================================
from tkinter import simpledialog, colorchooser
import shutil

NOTE_UI_SETTINGS_FILE = os.path.join(APP_DIR, "쪽지설정.json")
WORK_GROUP_NAME = "!!!사용자그룹추가"
APP_ICON_PNG_B64 = "iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAAOxUlEQVR4XtVbC3RU1bn+9j5nZpJMEpIAiS20lVsfRUUEFEVEsRbow+WDmtXeq7VoirS3tS567+1qu7qK8Vpd1aq12t5K0VLbakvaAm1VyhUFUnRpLxBIpUKMyCvvkMdMMjPnnL33Xf8+M5MHycyZ5ITSb61ZQIacs//v//5///vfezPkBobKSo62CxiizSy3X51gFL5PofyAQs0GCTDl89sqDVRWGj4/dOJAY/U43mxeZMAaBlRLPdpZXyo1uJytDHUeJCqUEmcEKYxxCcbaGNAgOPZh99oOl901HLhXZVJEBgLI4zVCP2fuFxYxadzFmFoCxivAknZno+90ISV4JQApOxRn27hSPxV7127TQyA11CRtGYaRTUj9wsW3lTMz/AgDu00bLW1ACQkFVxFnGhg4GOfgAUBJKCl/p2xrNd5af2w0Ek4lIGX8JSsu4yy4ATxwNkScjFakNeCM8ftooJFKQDGYeRzSbpG2/Vn87Wc7RiJhGAEUM9USs74wj5vGK2CsGMJ2wJh5pjnbE5RywE0ae1w69jLUP7NzOAmDCCDjAcxrLuNC7QE3PvBPbXwKSglw04BSHVLIOaj/6Qng3nRid43WieKA/iFzxA9hhnIynjEGg3OYxun50LvonZ7AmAHpODACUxjHU3pGWDPwm+5TUnE/584FnAVfg3AEGLJOcTQQgm3ZQMIGxGnKjQYHgibMUFAbIKSH9yoImAGDOWKZ2Ld2a8rmIbHNJP8KTLLbprlzVIKJfM457EhMZ9uzppfjwg+/H+8vL9UqmLA0qQBHSrS09+BAYxNOHGvVYzSLCqCkglKZCkD6kikJeQ+ArSnjyEr6KMy7axJ35LvgvEw/bRQztPSUgojEcNWVF+Ke25bguitmorQ47EmRfqE3GsP2v76NJ371Ml7esQ88nAdm8OTQRwQRwKBkvzSNc7B7bTMVSlTb6yLBmH3Xtcrgr0DaVEsP5IZB0MaTzIXAQ1//LL52+7L0tzLpAb+L8JHAGQPnAwp9qmY77vnuL2FLBRYwRidBh0HQYELcJOqe2oxrrjFNvbAhCXB5PlhAQdlSFxTDQLYzpaCEwIbH78by6+ZCSqmFRYNxB+QxMfkAkjuRTq9cVbkYM6ZPxY1fegwOOYjzUcJBUS0DBWcmgM2Ins94elWnZAXY6BZQzDuRfjz8zVu18Zbt6J8ZRg4Z2QeQwshwPfMYXCdiGsvSBRfix9V3wOmLZ54hFMU7KvRQCt+n0p5WavR5hV5CCW/BlRdh9a0fgyMEAoHTWxuRxwUZDldxgxEMmHAcgTtuXIilS+ZpR6VmqJGzwYCtI8b6KaD3SYmvr/h48qsMUvER5G0yOu3xpOF7D7Vorw/PT0TSN+74BMBJ5t6QlQBKOHbCRsUHy3Hd5RQ6yMyuD5CDvE1Gk8ePt0Ww/sW/44qVP8cjv9gG0zA0OSnwZCguvOQczDhnGpy4pceeDVl1rDN/wsasc6ajKJynvTFcgn6AkikZrmsMnXEByxao3XcCG2sbsWt/E460dKDrRAO+v/5uPQa3AHKN1AWRkDoc5p7/ARx++xhYfhAQmbXggQCtQ0wrL9EvokFyH7O9nj61qhiMpMcajnfjT7vexQuvH0bDsW79TulE0dfdhMtmnYWFc87Tch8e5+6TgGkVpTpkvQRqVgLSEhuWeMaDlLcHz+eRfguv7jmOTTvfwetvNaM7mkA4L4hwvomurjYkrBiszjbc/tVlWpWUiCkMRkLGBDgMngnwAyN5e39jBzbXNmLrm0dwuLlHGxfOMzG1NIx4PI6W1mYIJRGPdGHSlDA+8/HL9VBoCvYDE06ALliUK/GUijp6YvjfN49i018asedgK/riNgpCAZQUhpI2cUSivTjZ1ebGt3Jgtbfg1s8tw9TSQh3rVAP4gQkhIF2sJOdsQy8fFN440IJNOxuxbfdRNHX06YUTSbysOF9XlZTVOePojZ5ET29nsgGl0NfbAZYXxKpbrnZtzp7cPcNXAihZUXzrhJbMGcfbo3jp9cP44653Uf9uByxbIpwfQGlxyM0FUkEIkTa2s6sVff0RLXHqwCX6OuD0RnDZ/JmYf9GMEZPfeOArAUOmr/1N2LjzHdTWnUB7dz9CQUPLPJznziQiPT2R1w04wtbGJxLxtPFWPAJhxfRUdueNC7Mmv7HAVwION/dic+07+NNrh3HoWJf2Lnm7bFK+XqGR4cNLNM4NxBMxnOxqhSPc9QUlQuFYsPq7dTOnbNoUVC65VNuXS4b3gnETQEY5jsTDz72BxzfshWIBTC4pQElhHoQUSYmP3LEhY6N9veju6dDSpvgn0N/j0U6dO0RfApWV12Jyib/JL4Vxp1LylmlyfPGm2Xho1aWYMTmB1qaj6I32wzSD+pNqorgdG3clR5/uns6BTJ8sgnTc93dDOAldxFHyq1p+le/Jzz8CKNMzhsmTwlhxw+V45X+qcP/K+SgUJ/DW/jfR1nYcDBJmIAjDCGjJU8bvONmC3kjXkPmcjHesflixCEzDhOiPY/7cc3HpBTO00vyWv64nvMSJF+iVm5AoLCjAl//1Y3jj+W/hsa99AvlOCw68+SoaG+rQF22HZcXQ3tmEWKxPkzEABikdxKMn3eYLzSKWg6qbrtL/JtImAr4RQEqg4oRUTl2ZcH4Id9+6FPWbH8Cj3/43lBc4OLL3NbxzaK9upBq6jB1YqJCRZDyRQEqw4xamfrB8wpKfbyEwHGQIFThEBCmiqCAPq29fhvo/PIhHq6swvSQPJ9uowhtYqpDBJHuSP/2dpK764qhcehlKigv0czzvA+QIzwrwsLQeAvr/riKUVkRRQQirP78M+zffjwfu/hSkFO4qkJQgbVgxWge4w6FlLi/IQ9XNY0t+uYzVMwGW5e6U5wo9Swwigtrn36xahhXXz0NvJIZ4IoGepqMQlg3TNPT/FX1xXDHvPMyd+aGckx+RmrCGdosyISsBeuYyDRxp7tTPSZW4uSJFBHmXFPHtFdeiuDAf/71qCR76xmdQXlIEq7PX7fJIiZWfXqRfkUvyozqCctF7TZ000HR/IBOyFkJSSbBQAPsbjqO9K4IppUVuUyQXnQ2Cjm8FlJcV4sUfrMCVsz6kJXv7DQvx4+e24Qc//zPU1BIs/+jc5Loiq480Uu35aH8Cu/9+BMgLZtok8Z4E6cFmMICe5pPYvL0u2R8d35Tkbi4pLLzYNZ4UUVFWjOqv3Iy6Tffj19+7S6sjVTR5QWp/cMuuerS814pAKOCW3lngiV5F7aW8AB565iUdXyTnzPtw2UHP0BsbOsLcZEnZfsa0KfjkoovdnTyPKkuNhQx+YN0LenfIjd3s8EQAPdjMD6HhwBH856MbtIxpsF4YzpiABrXZUhsdbkeYpr3sg3el784yROJ3frQJe//vEMwwrUN8JIBABgdKCvHk0y/iwXUv6IxNeYA2JFKe9AP0zGwlb6qPQH1BIi5gGnjy+W347hMbEZgUhsxhmz5rEhwM8kygqADf+t7zaDzahgdX34KppUUj2q3ncq9uzAV6IeWGB4eB7kg/vvPERjyxfov2PGUnb753kRMBWglKIVAcxtPPvYwtf6lH1fJFWL70Ulx87nS3utN1+8QsXFKIJSw0HGnFH7fXYd1vd+K9xiaYJeFkvyEX88dAgCaBlFBSiBOdvbjvvmdR39iE3z/2ZXdnFkq3q9f9vhZrn9sGIz/oW4i4vUaJjq4ojrV0Qvb2AwUhBJKN0rEgZwWkQDkhFDThlBahcqnbrUn97NFfbMV/3Pes3qb2s4GZhmGAB00Ey4oyNly8YMwEkNQty0bh5GJcO/8jOhsTIdU/+QPuffg3OhnpTUra7vKRBb37k9xGcxU3PoyZACo7RSyOBQsuwlmTJ+lR/NcjG/D9H21CIFkt0qFSF7nF5enEmAlINSxuvm6OHu/KNeux7tk/Izi5WOcIz7nI43zvCWPgeUwEkPzpaFxhRQmuv3o2Vt67HuvWb0GwvETXBV5ABEqHTrS6Z5fGDQo1zrIdlDoFaQIY837JQJfCCRszL5qB1Q//Gr+t2ZGT8Tph9lswi/OShzjHudtDHHJAxB3dReb5dFg6A6WDbDX1TQsN1u5VQvqMQH4Ie98+ir/ubYA5ZZJnz+uMQNvty+dg8oJ/8S8CaG1hC7RseQvttQ3gISJhBBb00QPVrr+JNjNTXzMh8wUOKS7oW0/rA+1F4a4RvK7ZSaKiz8KUq8/F9OVz4ETiSe7HmSeTZxfpnODZn7sCsaYeRBvbwUPmMBL0Kg5MsoP6xYUHFUdNjU7VQsg9kE63forHtJ2q+nKCUig6rxwi4UBYwr1+4Ej3z7F+HPc5Tr+lp97CD08FbFonDBkZbUgYEFZcKPWG/mbHYukaS4clDzx9Uin1MrhJO5zeMtlYwBiihztghExwWrZSseTLh8HIC+gR9R3pBEw+1I10h8AwqSypRf264+51mmo5ZBYwGH9SKnWLP2n5VND+ICWozl2NyD9rEsrmnz0WCkcGNVkcG60b6xA52AKeFxjWs3CnGy7lk9q7dDq+ZnChmjwyy2bfuYmZ+TfCiU/cRQk6NGEJBMvCOTUvMkFvqMYs2D0xbfxQ8SsHZshUIvGqqnv6o1izhqPavQg2oICaC6jGZIpX/TuT1gJwsxySjs2nbkj5CFrz5wdg99Jpcx+fS2FQMKwXSOeDuUF3BnqVYazUE271wDuHponUvYFZKxZxM0RHyvP0/vREXZnRb/ehCEpj2Pa7vjJjmKC8r6zrRd3PXspwZSaJFAmzqxZzbvwGPFAOJ+EmRXeK9HPEE4HUpSnADBlQopsJ+zax75kXhhs/Ov1pEj5/NmOhx5lh3KDtFnRtjmYIz5X+aYZuFxkgp1O5I8VW5dhfRf0zB+loPHbsOGXHZHRvDrpsaFyy6pMSahWDWgwjUHzmioDWyXafUryWc7FW7Fm3UXsg54uTaSRvkqWuzs5eNc3gbJ6C+IhSqIBS1H/+B0ZF8t0MkoG3Mc4OCqZ2Y/fao0kThl79HTP0ReQUGf8MWMP9ujw9DGs4rtnOgcVnKAvbdXmbVqwH/D/40QlZaWMSZwAAAABJRU5ErkJggg=="


def log_dir():
    d = os.path.join(APP_DIR, "logs")
    try:
        os.makedirs(d, exist_ok=True)
    except Exception:
        pass
    return d


def open_run_log(kind):
    path = os.path.join(log_dir(), f"{datetime.now():%Y%m%d_%H%M%S}_{kind}.txt")
    try:
        f = open(path, "w", encoding="utf-8")
        f.write(f"메신저 작업 도우미 {APP_VERSION} 작업 기록\n")
        return f
    except Exception:
        return open(os.devnull, "w", encoding="utf-8")


def guess_kw_columns(headers, kw_vars, kw_cbs):
    hs = list(headers)
    for cb in kw_cbs:
        cb.configure(values=[""] + hs)
    rules = [
        (["검색어1", "수신자명", "수신자", "대상자명", "성명", "이름"], ["수신자", "성명", "이름", "대상자"]),
        (["검색어2", "부서/기관명", "부서", "기관명", "소속"], ["부서", "기관", "소속", "학교"]),
        (["검색어3", "직급", "직위"], ["직급", "직위"]),
    ]
    chosen = []
    for i, (exact, contains) in enumerate(rules):
        cur = kw_vars[i].get()
        if cur in hs and cur not in chosen:
            chosen.append(cur)
            continue
        g = guess_column([h for h in hs if h not in chosen], exact, contains)
        if i == 0 and not g and hs:
            g = hs[0]
        kw_vars[i].set(g)
        if g:
            chosen.append(g)


class ColumnPickDialog(tk.Toplevel):
    """엑셀을 열 때 검색어1~3 열과 대상 고르기를 정하는 창."""

    def __init__(self, parent, headers, rows):
        super().__init__(parent)
        self.title("검색어 열 고르기")
        self.resizable(False, False)
        self.result = None
        self.headers, self.rows = headers, rows
        frm = ttk.Frame(self, padding=14)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="검색어1·2로 찾고, 여러 명이면 검색어3으로 한 번 더 좁힙니다.").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))
        self.kw_vars, self.kw_cbs = [], []
        for i in range(3):
            ttk.Label(frm, text=f"검색어{i + 1}").grid(row=1 + i, column=0, sticky="w", pady=2)
            v = tk.StringVar()
            cb = ttk.Combobox(frm, textvariable=v, state="readonly", width=24)
            cb.grid(row=1 + i, column=1, sticky="w", pady=2)
            self.kw_vars.append(v)
            self.kw_cbs.append(cb)
        guess_kw_columns(headers, self.kw_vars, self.kw_cbs)
        ttk.Label(frm, text="대상 고르기").grid(row=4, column=0, sticky="w", pady=(8, 2))
        self.fcol = tk.StringVar(value="구분" if "구분" in headers else "(전체)")
        self.fval = tk.StringVar(value="(모두)")
        fc = ttk.Combobox(frm, textvariable=self.fcol, state="readonly", width=24, values=["(전체)"] + list(headers))
        fc.grid(row=4, column=1, sticky="w", pady=(8, 2))
        self.fv_cb = ttk.Combobox(frm, textvariable=self.fval, state="readonly", width=18)
        self.fv_cb.grid(row=4, column=2, sticky="w", padx=(6, 0), pady=(8, 2))
        fc.bind("<<ComboboxSelected>>", lambda e: self._fill_vals())
        self._fill_vals()
        bar = ttk.Frame(frm)
        bar.grid(row=5, column=0, columnspan=3, sticky="e", pady=(12, 0))
        ttk.Button(bar, text="불러오기", command=self._ok).pack(side="left", padx=4)
        ttk.Button(bar, text="취소", command=self.destroy).pack(side="left")
        self.transient(parent)
        self.grab_set()
        self.wait_window(self)

    def _fill_vals(self):
        col = self.fcol.get()
        if col in ("", "(전체)"):
            self.fv_cb.configure(values=[])
            self.fval.set("")
            return
        vals = []
        for r in self.rows:
            v = r["cells"].get(col, "")
            if v not in vals:
                vals.append(v)
        self.fv_cb.configure(values=["(모두)"] + vals)
        self.fval.set("(모두)")

    def _ok(self):
        if not self.kw_vars[0].get():
            messagebox.showwarning("검색어", "검색어1 열을 고르세요.", parent=self)
            return
        col = self.fcol.get()
        val = self.fval.get()
        fcol = None if col in ("", "(전체)") or val in ("", "(모두)") else col
        self.result = ([v.get() for v in self.kw_vars], fcol, val if fcol else None)
        self.destroy()


class FastGrid(ExcelLikeGrid):
    """엑셀처럼 입력되는 표.
    - 선택한 칸에 입력칸이 늘 대기 → 한글(IME)이 첫 글자부터 끊기지 않음
    - 바로 치면 덮어쓰기, F2·더블클릭은 고치기(커서 이동 가능)
    - Enter/Tab/방향키 이동, Shift+방향키 범위 선택, Ctrl+방향키 끝으로, Home/End/PageUp/PageDown
    - Ctrl+C/V, Delete, BackSpace(지우고 입력), Esc(취소), Ctrl+Z(되돌리기)
    - 보이는 행만 그려 1,000행 이상에서도 빠름
    """

    SEL_BG = "#d9ecff"

    def _build(self):
        super()._build()
        self._last_yview = None
        self.body_canvas.configure(yscrollcommand=self._on_yscroll)
        self._undo = []
        self.mode = "ready"            # ready: 칸 선택 / enter: 바로 쳐서 덮어쓰기 / edit: F2로 고치기
        self._edit_row, self._edit_key, self._edit_original = 0, None, ""
        self._suppress = False
        self.in_var = tk.StringVar()
        self.input = tk.Entry(self.body_canvas, textvariable=self.in_var, font=("Malgun Gothic", 10),
                              relief="flat", bd=0, highlightthickness=0, insertwidth=1, bg=self.SEL_BG)
        self.input_win = None
        self.in_var.trace_add("write", self._on_input_change)
        b = self.input.bind
        b("<KeyPress>", self._on_input_key)
        for seq in ("<Control-c>", "<Control-C>"):
            b(seq, self._input_copy)
        for seq in ("<Control-v>", "<Control-V>"):
            b(seq, self._input_paste)
        for seq in ("<Control-a>", "<Control-A>"):
            b(seq, lambda e: self._select_all() if self.mode == "ready" else None)
        for seq in ("<Control-z>", "<Control-Z>"):
            b(seq, self._undo_last)
        b("<MouseWheel>", self._on_mousewheel)
        b("<Shift-MouseWheel>", self._on_shift_mousewheel)
        b("<FocusOut>", self._on_input_focus_out)
        self.body_canvas.bind("<FocusIn>", lambda e: self.input.focus_set())

    # ---------- 그리기 ----------
    def _on_yscroll(self, first, last):
        self.vsb.set(first, last)
        cur = (round(float(first), 5), round(float(last), 5))
        if cur != self._last_yview:
            self._last_yview = cur
            self.after_idle(self._draw_body)

    def _draw_body(self):
        c = self.body_canvas
        c.delete("all")
        self.input_win = None
        total_w = self._total_width()
        total_h = max(1, len(self.data) * self.row_height)
        c.configure(scrollregion=(0, 0, total_w, total_h))
        widths = self._col_widths()
        try:
            top = int(c.canvasy(0))
            height = max(c.winfo_height(), self.row_height * 20)
        except Exception:
            top, height = 0, self.row_height * 40
        r_start = max(0, top // self.row_height - 2)
        r_end = min(len(self.data), (top + height) // self.row_height + 3)
        for r in range(r_start, r_end):
            row = self.data[r]
            y1 = r * self.row_height
            y2 = y1 + self.row_height
            row_bg = self._row_bg(r)
            row_num_bg = self.SEL_BG if r == self.active_row else "#f4f7fb"
            c.create_rectangle(0, y1, self.row_header_width, y2, fill=row_num_bg, outline="#111111")
            c.create_text(self.row_header_width / 2, y1 + self.row_height / 2, text=str(r + 1), font=("Malgun Gothic", 10), fill="#000000")
            for col_idx, col in enumerate(self.columns_def):
                key = col["key"]
                x1 = self._col_left(col_idx)
                x2 = x1 + widths[col_idx]
                fill = self.SEL_BG if self._is_selected(r, col_idx) else row_bg
                c.create_rectangle(x1, y1, x2, y2, fill=fill, outline="#111111")
                value = self._fit_text(row.get(key, ""), widths[col_idx])
                anchor = "center" if key in ("status", "rowno") else "w"
                x_text = (x1 + x2) / 2 if anchor == "center" else x1 + 6
                c.create_text(x_text, y1 + self.row_height / 2, text=value, font=("Malgun Gothic", 10), fill="#000000", anchor=anchor)
        if self.data:
            r1, col1, r2, col2 = self._normalized_selection()
            x1, y1, _, _ = self._cell_bbox(r1, col1)
            _, _, x2, y2 = self._cell_bbox(r2, col2)
            c.create_rectangle(x1 + 1, y1 + 1, x2 - 1, y2 - 1, outline="#2b6cb0", width=2)
            if 0 <= self.active_row < len(self.data) and 0 <= self.active_col < len(self.columns_def):
                x1, y1, x2, y2 = self._cell_bbox(self.active_row, self.active_col)
                c.create_rectangle(x1 + 1, y1 + 1, x2 - 1, y2 - 1, outline="#1a73e8", width=2)
        self._place_input()

    def _place_input(self):
        c = self.body_canvas
        if self.input_win is not None:
            try:
                c.delete(self.input_win)
            except Exception:
                pass
            self.input_win = None
        if not self.data or not self.columns_def:
            return
        if self.mode == "ready":
            r, ci = self.active_row, self.active_col
        else:
            r, ci = self._edit_row, self._key_to_col(self._edit_key)
        r = max(0, min(r, len(self.data) - 1))
        x1, y1, x2, y2 = self._cell_bbox(r, ci)
        if self.mode == "ready":
            self.input.configure(bg=self.SEL_BG)
            self.input_win = c.create_window(x1 + 4, y1 + 4, window=self.input, anchor="nw", width=2, height=max(8, y2 - y1 - 8))
        else:
            self.input.configure(bg="#ffffff")
            self.input_win = c.create_window(x1 + 3, y1 + 3, window=self.input, anchor="nw",
                                             width=max(10, x2 - x1 - 5), height=max(10, y2 - y1 - 5))

    def _ensure_visible(self, r, ci=None):
        c = self.body_canvas
        try:
            total_h = max(1, len(self.data) * self.row_height)
            top = c.canvasy(0)
            h = max(self.row_height, c.winfo_height())
            y1 = r * self.row_height
            if y1 < top:
                c.yview_moveto(y1 / total_h)
            elif y1 + self.row_height > top + h:
                c.yview_moveto(max(0, (y1 + self.row_height - h)) / total_h)
            if ci is not None and c.winfo_width() > 60:   # 화면에 놓이기 전에는 가로 이동 안 함
                total_w = max(1, self._total_width())
                left = c.canvasx(0)
                w = max(50, c.winfo_width())
                cx1, cx2 = self._col_left(ci), self._col_right(ci)
                if cx1 < left + self.row_header_width:
                    self._xview("moveto", max(0, cx1 - self.row_header_width) / total_w)
                elif cx2 > left + w:
                    self._xview("moveto", (cx2 - w) / total_w)
        except Exception:
            pass

    # ---------- 칸 이동 ----------
    def focus_cell(self, row_idx, key_or_index=0):
        if not self.data:
            self.add_row()
        row_idx = max(0, min(int(row_idx), len(self.data) - 1))
        col_idx = key_or_index if isinstance(key_or_index, int) else self._key_to_col(key_or_index)
        col_idx = max(0, min(col_idx, len(self.columns_def) - 1))
        self.active_row, self.active_col = row_idx, col_idx
        self.selected_row = row_idx
        self.selected_col = self.columns_def[col_idx]["key"]
        self.sel_anchor = (row_idx, col_idx)
        self.selection = (row_idx, col_idx, row_idx, col_idx)
        self._ensure_visible(row_idx, col_idx)
        self._redraw()
        self.input.focus_set()

    def scroll_to_row(self, row_idx):
        self._ensure_visible(row_idx)

    def _last_data_row(self):
        for i in range(len(self.data) - 1, -1, -1):
            if any(safe_text(v) for v in self.data[i].values()):
                return i
        return 0

    def _nav(self, ks, shift=False, ctrl=False):
        r, c = self.active_row, self.active_col
        ncol = len(self.columns_def)
        page = max(1, self.body_canvas.winfo_height() // self.row_height - 1)
        if ks == "Up":
            r = 0 if ctrl else r - 1
        elif ks == "Down":
            r = self._last_data_row() if ctrl else r + 1
        elif ks == "Left":
            c = 0 if ctrl else c - 1
        elif ks == "Right":
            c = ncol - 1 if ctrl else c + 1
        elif ks == "Return":
            r = r - 1 if shift else r + 1
            shift = False
        elif ks == "Tab":
            c = c - 1 if shift else c + 1
            shift = False
        elif ks == "ISO_Left_Tab":
            c -= 1
            shift = False
        elif ks == "Prior":
            r -= page
        elif ks == "Next":
            r += page
        elif ks == "Home":
            c = 0
            if ctrl:
                r = 0
        elif ks == "End":
            c = ncol - 1
            if ctrl:
                r = self._last_data_row()
        c = max(0, min(c, ncol - 1))
        r = max(0, r)
        if r >= len(self.data):
            if ks in ("Down", "Return") and not ctrl:
                self.add_row()
                r = len(self.data) - 1
            else:
                r = len(self.data) - 1
        if shift:
            ar, ac = self.sel_anchor
            self.active_row, self.active_col = r, c
            self.selected_row, self.selected_col = r, self.columns_def[c]["key"]
            self.selection = (ar, ac, r, c)
            self._ensure_visible(r, c)
            self._redraw()
        else:
            self.focus_cell(r, c)
        return "break"

    def _move_key(self, key, shift=False):
        return self._nav(key, shift=shift)

    # ---------- 입력 ----------
    def _on_input_change(self, *args):
        if self._suppress or self.mode != "ready":
            return
        if not self.in_var.get():
            return
        col = self.columns_def[self.active_col] if self.columns_def else None
        if not col or not col.get("editable", True):
            self._set_input("")
            return
        self.mode = "enter"
        self._edit_row, self._edit_key = self.active_row, col["key"]
        self._edit_original = self.get_value(self.active_row, col["key"])
        self.selection = (self.active_row, self.active_col, self.active_row, self.active_col)
        self.sel_anchor = (self.active_row, self.active_col)
        self._redraw()
        self.input.icursor("end")

    def _set_input(self, text):
        self._suppress = True
        try:
            self.in_var.set(text)
        finally:
            self._suppress = False

    def _on_input_key(self, event):
        ks = event.keysym
        shift = bool(event.state & 0x0001)
        ctrl = bool(event.state & 0x0004)
        if self.mode == "ready":
            if ks in ("Up", "Down", "Left", "Right", "Tab", "ISO_Left_Tab", "Return", "Prior", "Next", "Home", "End"):
                return self._nav(ks, shift, ctrl)
            if ks == "F2":
                self._begin_edit()
                return "break"
            if ks == "Delete":
                self._clear_selection()
                return "break"
            if ks == "BackSpace":
                self._begin_edit(initial="")
                return "break"
            if ks == "Escape":
                return "break"
            return None          # 글자는 입력칸이 그대로 받음(한글 조합 포함)
        # 입력 중
        if ks == "Escape":
            self._cancel_edit()
            return "break"
        move = None
        if ks == "Return":
            move = "Up" if shift else "Return"
        elif ks == "Tab":
            move = "ISO_Left_Tab" if shift else "Tab"
        elif ks == "ISO_Left_Tab":
            move = "ISO_Left_Tab"
        elif ks in ("Up", "Down"):
            move = ks
        elif ks in ("Left", "Right") and self.mode == "enter":
            move = ks
        if move:
            # 한글 마지막 글자 조합이 먼저 들어오도록 아주 잠깐 기다렸다가 저장
            self.after(40, lambda m=move: self._commit_and_move(m))
            return "break"
        return None

    def _begin_edit(self, initial=None):
        if not self.data or not self.columns_def:
            return
        col = self.columns_def[self.active_col]
        if not col.get("editable", True):
            return
        self._commit_edit()
        self._edit_row, self._edit_key = self.active_row, col["key"]
        self._edit_original = self.get_value(self.active_row, col["key"])
        self.mode = "edit" if initial is None else "enter"
        self._set_input(self._edit_original if initial is None else initial)
        self._redraw()
        self.input.focus_set()
        self.input.icursor("end")
        self.input.selection_clear()

    def _commit_edit(self):
        if getattr(self, "mode", "ready") == "ready":
            return
        val = self.in_var.get()
        r, k = self._edit_row, self._edit_key
        self.mode = "ready"
        self._set_input("")
        if k is not None and val != self._edit_original:
            self._push_undo([(r, k, self._edit_original)])
            self.set_value(r, k, val, redraw=False)
            self.clear_result(r, redraw=False)
        self._redraw()

    def _cancel_edit(self):
        if getattr(self, "mode", "ready") == "ready":
            return "break"
        self.mode = "ready"
        self._set_input("")
        self._redraw()
        self.input.focus_set()
        return "break"

    def _commit_and_move(self, direction):
        if self.mode != "ready":
            self._commit_edit()
        return self._nav(direction)

    def _on_input_focus_out(self, event=None):
        if self.mode != "ready":
            self.after(40, self._commit_edit)

    # ---------- 복사·붙여넣기·지우기·되돌리기 ----------
    def _input_copy(self, event=None):
        if self.mode != "ready":
            return None
        return self._copy_selection()

    def _input_paste(self, event=None):
        if self.mode == "ready":
            return self._paste_clipboard()
        try:
            text = self.clipboard_get()
        except Exception:
            return None
        if clipboard_text_has_table_shape(text):
            row, key = self._edit_row, self._edit_key
            self._cancel_edit()
            self.focus_cell(row, key)
            return self._paste_clipboard()
        return None

    def _push_undo(self, changes):
        if changes:
            self._undo.append(changes)
            del self._undo[:-50]

    def _undo_last(self, event=None):
        if self.mode != "ready":
            return None
        if not self._undo:
            return "break"
        changes = self._undo.pop()
        for r, k, old in changes:
            if 0 <= r < len(self.data):
                self.data[r][k] = old
        r0 = changes[0][0]
        self.focus_cell(min(r0, len(self.data) - 1), changes[0][1])
        return "break"

    def paste_matrix(self, start_row, start_key, matrix):
        keys = self.editable_keys()
        if matrix and keys:
            sk = start_key if start_key in keys else keys[0]
            sc = keys.index(sk)
            old = []
            for ro, vals in enumerate(matrix):
                for co, _ in enumerate(vals):
                    if sc + co < len(keys) and start_row + ro < len(self.data):
                        old.append((start_row + ro, keys[sc + co], self.data[start_row + ro].get(keys[sc + co], "")))
            n_before = len(self.data)
            super().paste_matrix(start_row, start_key, matrix)
            for ro, vals in enumerate(matrix):          # 새로 생긴 행은 빈 값으로 되돌림
                for co, _ in enumerate(vals):
                    if sc + co < len(keys) and start_row + ro >= n_before:
                        old.append((start_row + ro, keys[sc + co], ""))
            self._push_undo(old)
        else:
            super().paste_matrix(start_row, start_key, matrix)
        self.input.focus_set()

    def _clear_selection(self):
        r1, c1, r2, c2 = self._normalized_selection()
        old = []
        for r in range(r1, r2 + 1):
            for ci in range(c1, c2 + 1):
                col = self.columns_def[ci]
                if col.get("editable", True) and r < len(self.data):
                    old.append((r, col["key"], self.data[r].get(col["key"], "")))
        self._push_undo(old)
        super()._clear_selection()

    # ---------- 마우스 ----------
    def _on_button_1(self, event):
        if event.state & 0x0001 and self.data:      # Shift+클릭: 범위 선택
            self._commit_edit()
            cell = self._cell_at(event.x, event.y)
            if cell:
                r, c = cell
                ar, ac = self.sel_anchor
                self.active_row, self.active_col = r, c
                self.selection = (ar, ac, r, c)
                self._redraw()
            self.input.focus_set()
            return "break"
        res = super()._on_button_1(event)
        self.input.focus_set()
        return res

    def _on_drag(self, event):
        res = super()._on_drag(event)
        try:
            h = self.body_canvas.winfo_height()
            if event.y > h:
                self.body_canvas.yview_scroll(1, "units")
            elif event.y < 0:
                self.body_canvas.yview_scroll(-1, "units")
        except Exception:
            pass
        return res

    def _on_double_click(self, event):
        cell = self._cell_at(event.x, event.y)
        if cell:
            r, c = cell
            self.focus_cell(r, c)
            self._begin_edit()
        return "break"

    def set_rows(self, rows):
        self._commit_edit()
        self._undo = []
        self.data = []
        for values in rows:
            row = {col["key"]: "" for col in self.columns_def}
            for k, v in (values or {}).items():
                if k in row:
                    row[k] = safe_text(v)
            self.data.append(row)
        if not self.data:
            self.data.append({col["key"]: "" for col in self.columns_def})
        self.active_row = self.active_col = self.selected_row = 0
        self.sel_anchor = (0, 0)
        self.selection = (0, 0, 0, 0)
        self.body_canvas.yview_moveto(0)
        self._redraw()


class ZonedGrid(FastGrid):
    """구역이 나뉜 수신자 표.
    - 제목 위에 구역 띠: 검색어 / 개별문구·첨부파일 / 실행 결과
    - 개별문구 열 제목을 누르면 이름 바꾸기, 제목에서 우클릭하면 열 추가·삭제
    """

    ZONES = {"search": ("검색어 (사람 찾기)", "#dbe7f3", "#c5d9ee"),
             "field": ("개별문구 · 첨부파일 (쪽지 내용)", "#e2efda", "#cfe5c2"),
             "result": ("실행 결과", "#fce4d6", "#f8cfb6")}
    BAND_H = 22

    def __init__(self, parent, columns, rows=12, on_paste=None, on_add=None, on_delete=None, on_rename=None):
        self.on_add, self.on_delete, self.on_rename = on_add, on_delete, on_rename
        self._hdr_edit = None
        super().__init__(parent, columns, rows=rows, on_paste=on_paste)

    def _build(self):
        self.header_height = 28 + self.BAND_H
        super()._build()
        self.header_canvas.configure(height=self.header_height)
        self.header_canvas.bind("<Button-3>", self._on_header_right)

    def _col_at_x(self, event_x):
        x = self.header_canvas.canvasx(event_x)
        for i in range(len(self.columns_def)):
            if self._col_left(i) <= x < self._col_right(i):
                return i
        return None

    def _draw_header(self):
        c = self.header_canvas
        c.delete("all")
        c.configure(scrollregion=(0, 0, self._total_width(), self.header_height))
        bh = self.BAND_H
        c.create_rectangle(0, 0, self.row_header_width, self.header_height, fill="#eef2f7", outline="#111111")
        c.create_text(self.row_header_width / 2, bh + (self.header_height - bh) / 2, text="번호",
                      font=("Malgun Gothic", 10, "bold"), fill="#17324d")
        # 구역 띠
        i = 0
        cols = self.columns_def
        while i < len(cols):
            z = cols[i].get("zone", "search")
            j = i
            while j + 1 < len(cols) and cols[j + 1].get("zone", "search") == z:
                j += 1
            label, fill, band = self.ZONES.get(z, ("", "#dbe7f3", "#c5d9ee"))
            x1, x2 = self._col_left(i), self._col_right(j)
            c.create_rectangle(x1, 0, x2, bh, fill=band, outline="#111111", width=1)
            c.create_text((x1 + x2) / 2, bh / 2, text=label, font=("Malgun Gothic", 9, "bold"), fill="#17324d")
            i = j + 1
        for idx, col in enumerate(cols):
            x1, x2 = self._col_left(idx), self._col_right(idx)
            _, fill, _ = self.ZONES.get(col.get("zone", "search"), ("", "#dbe7f3", ""))
            c.create_rectangle(x1, bh, x2, self.header_height, fill=fill, outline="#111111")
            title = col.get("header", col["key"])
            if col.get("kind") == "field":
                title += " ✎"
            c.create_text((x1 + x2) / 2, bh + (self.header_height - bh) / 2, text=title,
                          font=("Malgun Gothic", 10, "bold"), fill="#17324d")

    # ---------- 제목 누르기: 이름 바꾸기 ----------
    def _on_header_click(self, event):
        idx = self._col_at_x(event.x)
        if idx is not None and self.columns_def[idx].get("kind") == "field" \
                and self.header_canvas.canvasy(event.y) >= self.BAND_H:
            self.start_header_rename(self.columns_def[idx]["key"])
            return "break"
        return super()._on_header_click(event)

    def start_header_rename(self, key):
        self._end_header_rename(commit=False)
        idx = next((i for i, c in enumerate(self.columns_def) if c["key"] == key), None)
        if idx is None:
            return
        x1, x2 = self._col_left(idx), self._col_right(idx)
        var = tk.StringVar(value=self.columns_def[idx].get("header", ""))
        ent = tk.Entry(self.header_canvas, textvariable=var, font=("Malgun Gothic", 10, "bold"),
                       justify="center", relief="solid", bd=1)
        win = self.header_canvas.create_window(x1 + 1, self.BAND_H + 1, anchor="nw", window=ent,
                                               width=x2 - x1 - 2, height=self.header_height - self.BAND_H - 2)
        self._hdr_edit = (ent, win, key, var)
        ent.bind("<Return>", lambda e: (self._end_header_rename(commit=True), "break")[1])
        ent.bind("<KP_Enter>", lambda e: (self._end_header_rename(commit=True), "break")[1])
        ent.bind("<Escape>", lambda e: (self._end_header_rename(commit=False), "break")[1])
        ent.bind("<FocusOut>", lambda e: self.after(1, lambda: self._end_header_rename(commit=True)))
        ent.focus_set()
        ent.select_range(0, "end")

    def _end_header_rename(self, commit=True):
        if not self._hdr_edit:
            return
        ent, win, key, var = self._hdr_edit
        self._hdr_edit = None
        new = var.get().strip()
        try:
            self.header_canvas.delete(win)
            ent.destroy()
        except Exception:
            pass
        old = next((c.get("header", "") for c in self.columns_def if c["key"] == key), "")
        if commit and new and new != old and self.on_rename:
            self.after(1, lambda: self.on_rename(key, new))

    # ---------- 제목 우클릭: 열 추가·삭제 ----------
    def _on_header_right(self, event):
        self._end_header_rename(commit=True)
        idx = self._col_at_x(event.x)
        if idx is None:
            return "break"
        col = self.columns_def[idx]
        m = tk.Menu(self, tearoff=0, font=("Malgun Gothic", 10))
        if col.get("zone") != "field":
            m.add_command(label="이 구역은 열을 바꿀 수 없습니다", state="disabled")
        else:
            nfield = sum(1 for c in self.columns_def if c.get("kind") == "field")
            if col.get("kind") == "field":
                m.add_command(label="왼쪽에 열 추가", command=lambda: self.on_add and self.on_add(col["key"], "left"))
                m.add_command(label="오른쪽에 열 추가", command=lambda: self.on_add and self.on_add(col["key"], "right"))
                m.add_separator()
                m.add_command(label="열 이름 바꾸기", command=lambda: self.start_header_rename(col["key"]))
                m.add_command(label="열 삭제", state=("normal" if nfield > 1 else "disabled"),
                              command=lambda: self.on_delete and self.on_delete(col["key"]))
            else:   # 첨부파일 열
                m.add_command(label="왼쪽에 열 추가", command=lambda: self.on_add and self.on_add(col["key"], "left"))
        try:
            m.tk_popup(event.x_root, event.y_root)
        finally:
            m.grab_release()
        return "break"


class RichTemplateEditor(tk.Frame):
    """글자 크기·굵게·밑줄·색을 지정하는 쪽지 서식 편집기. 내용은 runs로 주고받는다."""

    COLORS = [("검정", "#000000"), ("파랑", "#0000FF"), ("빨강", "#FF0000"), ("남색", "#1B1760"),
              ("진남", "#000080"), ("갈색", "#803300"), ("초록", "#008000")]
    SIZES = [9, 10, 11, 12, 13, 14, 15, 16, 18, 20]

    def __init__(self, parent, height=12, readonly=False, on_change=None, toolbar=True):
        super().__init__(parent, bg="#ffffff")
        self.readonly = readonly
        self.on_change = on_change
        self._tags = set()
        self.files = {}          # 그림 이름 → 파일 경로
        self._imgs = {}          # 본문 속 그림(Tk 이름) → (그림 이름, PhotoImage)
        self._wins = {}          # 본문 속 첨부 표시(위젯 경로) → 자리표시 글자
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        if toolbar and not readonly:
            self._build_toolbar()
        wrap = tk.Frame(self, bg="#ffffff")
        wrap.grid(row=1, column=0, sticky="nsew")
        wrap.columnconfigure(0, weight=1)
        wrap.rowconfigure(0, weight=1)
        self.text = tk.Text(wrap, height=height, wrap="word", undo=True, font=("Malgun Gothic", 11),
                            relief="solid", bd=1, padx=8, pady=6)
        self.text.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.text.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.text.configure(yscrollcommand=sb.set)
        if readonly:
            self.text.configure(state="disabled", bg="#fbfcfe")
        else:
            self.text.bind("<KeyRelease>", self._after_edit)
            self.text.bind("<<Paste>>", self._on_paste)
            self.text.bind("<Control-b>", lambda e: (self.toggle("bold"), "break")[1])
            self.text.bind("<Control-u>", lambda e: (self.toggle("ul"), "break")[1])

    def _build_toolbar(self):
        bar = tk.Frame(self, bg="#ffffff")
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 4))
        tk.Label(bar, text="크기", bg="#ffffff", font=("Malgun Gothic", 9)).pack(side="left")
        self.size_var = tk.StringVar(value="12")
        cb = ttk.Combobox(bar, textvariable=self.size_var, values=[str(s) for s in self.SIZES], width=4, state="readonly")
        cb.pack(side="left", padx=(2, 6))
        cb.bind("<<ComboboxSelected>>", lambda e: self.apply("size", int(self.size_var.get())))
        tk.Button(bar, text="굵게", font=("Malgun Gothic", 9, "bold"), command=lambda: self.toggle("bold"), relief="groove", padx=6).pack(side="left", padx=1)
        tk.Button(bar, text="밑줄", font=("Malgun Gothic", 9, "underline"), command=lambda: self.toggle("ul"), relief="groove", padx=6).pack(side="left", padx=(1, 8))
        for name, color in self.COLORS:
            tk.Button(bar, text="■", fg=color, font=("Malgun Gothic", 11), bd=0, relief="flat", cursor="hand2",
                      command=lambda c=color: self.apply("color", c)).pack(side="left")
        tk.Button(bar, text="색…", font=("Malgun Gothic", 9), relief="groove", command=self._pick_color).pack(side="left", padx=(4, 0))
        tk.Label(bar, text="  글자를 블록으로 고른 뒤 누르기", bg="#ffffff", fg="#6b7280", font=("Malgun Gothic", 8)).pack(side="left")

    def _pick_color(self):
        c = colorchooser.askcolor(title="글자색")
        if c and c[1]:
            self.apply("color", c[1].upper())

    def _tag_for(self, st):
        st = norm_style(st)
        name = "st_%d_%d_%d_%s" % (st["size"], st["bold"], st["ul"], st["color"][1:])
        if name not in self._tags:
            font = ("Malgun Gothic", st["size"]) + (("bold",) if st["bold"] else ()) + (("underline",) if st["ul"] else ())
            self.text.tag_configure(name, font=font, foreground=st["color"])
            self._tags.add(name)
        return name

    @staticmethod
    def _style_of_tag(name):
        try:
            _, size, b, u, color = name.split("_")
            return {"size": int(size), "bold": b == "1", "ul": u == "1", "color": "#" + color}
        except Exception:
            return dict(DEFAULT_STYLE)

    def _style_at(self, index):
        for t in self.text.tag_names(index):
            if t.startswith("st_"):
                return self._style_of_tag(t)
        return dict(DEFAULT_STYLE)

    def _set_style(self, start, end, st):
        for t in self.text.tag_names(start):
            if t.startswith("st_"):
                self.text.tag_remove(t, start, end)
        self.text.tag_add(self._tag_for(st), start, end)

    @staticmethod
    def _load_photo(path, maxw=420):
        """편집기에 보일 그림. PIL이 있으면 어떤 형식이든, 없으면 png/gif만."""
        try:
            from PIL import Image, ImageTk
            im = Image.open(path)
            if im.width > maxw:
                im = im.resize((maxw, max(1, int(im.height * maxw / im.width))))
            return ImageTk.PhotoImage(im)
        except Exception:
            pass
        try:
            ph = tk.PhotoImage(file=path)
            n = max(1, -(-ph.width() // maxw))
            return ph.subsample(n) if n > 1 else ph
        except Exception:
            return None

    def _put_image(self, index, name):
        path = self.files.get(name)
        ph = self._load_photo(path) if path and os.path.exists(path) else None
        if ph is None:
            return False
        emb = self.text.image_create(index, image=ph, padx=2, pady=2)
        self._imgs[emb] = (name, ph)
        return True

    def _put_chip(self, index, token):
        """첨부 자리 표시(본문 안 작은 상자)를 넣는다. token: {첨부:이름} 또는 {개별첨부}."""
        if token == ROW_FILE_TOKEN:
            label, bg, fg = "[개별 첨부파일]", "#e8f2ff", "#1d4ed8"
        else:
            label, bg, fg = "[첨부] " + token[4:-1], "#fff4d6", "#8a5300"
        w = tk.Label(self.text, text=label, bg=bg, fg=fg, font=("Malgun Gothic", 9, "bold"),
                     bd=1, relief="solid", padx=4, cursor="arrow")
        self.text.window_create(index, window=w, padx=2, pady=1)
        self._wins[str(w)] = token
        return True

    def _put_token(self, index, m):
        tok = m.group(0)
        if tok.startswith("{그림:"):
            return self._put_image(index, tok[4:-1].strip())
        return self._put_chip(index, tok)

    def set_runs(self, runs, files=None):
        if files is not None:
            self.files = dict(files)
        if self.readonly:
            self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self._imgs = {}
        self._wins = {}
        for text, st in merge_runs(runs):
            pos = 0
            for m in BODY_TOKEN_RE.finditer(text):
                if m.start() > pos:
                    self.text.insert("end", text[pos:m.start()], (self._tag_for(st),))
                if not self._put_token("end", m):
                    self.text.insert("end", m.group(0), (self._tag_for(st),))
                pos = m.end()
            if pos < len(text):
                self.text.insert("end", text[pos:], (self._tag_for(st),))
        self.text.edit_reset()
        if self.readonly:
            self.text.configure(state="disabled")

    def get_runs(self):
        runs = []
        active = []
        for key, value, index in self.text.dump("1.0", "end-1c", text=True, tag=True, image=True, window=True):
            if key == "tagon" and value.startswith("st_"):
                active.append(value)
            elif key == "tagoff" and value in active:
                active.remove(value)
            elif key == "text":
                st = self._style_of_tag(active[-1]) if active else dict(DEFAULT_STYLE)
                runs.append([value, st])
            elif key == "image" and value in self._imgs:
                st = self._style_of_tag(active[-1]) if active else dict(DEFAULT_STYLE)
                runs.append(["{그림:" + self._imgs[value][0] + "}", st])
            elif key == "window" and value in self._wins:
                st = self._style_of_tag(active[-1]) if active else dict(DEFAULT_STYLE)
                runs.append([self._wins[value], st])
        return merge_runs(runs)

    def _on_paste(self, event=None):
        """한글·워드·웹에서 복사한 서식(HTML)을 그대로 붙인다. 서식 정보가 없으면 글자만."""
        html = read_clipboard_html()
        if html:
            try:
                prev = self.text.index("insert-1c") if self.text.compare("insert", ">", "1.0") else "insert"
                runs, _ = html_to_runs(html, self._style_at(prev))
                if runs_to_text(runs).strip("\x00"):
                    self.insert_runs(runs)
                    return "break"
            except Exception:
                traceback.print_exc()
        self.after(10, self._after_edit)
        return None

    def insert_runs(self, runs):
        """커서 자리에 서식 글자(와 그림)를 넣는다. 선택한 글자가 있으면 바꾼다."""
        try:
            self.text.delete("sel.first", "sel.last")
        except tk.TclError:
            pass
        for text, st in runs:
            pos = 0
            for m in re.finditer(r"\x00IMG\x00([^\x00]*)\x00", text):
                if m.start() > pos:
                    self.text.insert("insert", text[pos:m.start()], (self._tag_for(st),))
                self._paste_image_file(m.group(1))
                pos = m.end()
            if pos < len(text):
                self.text.insert("insert", text[pos:], (self._tag_for(st),))
        self.text.see("insert")
        self._changed()

    def _paste_image_file(self, src):
        """붙여넣은 서식 속 그림(한글이 임시 폴더에 만든 파일)을 쪽지첨부 폴더에 복사해 넣는다."""
        try:
            if not src or not os.path.exists(src):
                return False
            os.makedirs(NOTE_ATTACH_DIR, exist_ok=True)
            base = os.path.basename(src)
            stem, ext = os.path.splitext(base)
            name, n = base, 2
            while name in self.files or (os.path.exists(os.path.join(NOTE_ATTACH_DIR, name))
                                         and open(os.path.join(NOTE_ATTACH_DIR, name), "rb").read() != open(src, "rb").read()):
                name = f"{stem}({n}){ext}"
                n += 1
            dest = os.path.join(NOTE_ATTACH_DIR, name)
            if not os.path.exists(dest):
                shutil.copyfile(src, dest)
            self.files[name] = dest
            return self._put_image("insert", name)
        except Exception:
            traceback.print_exc()
            return False

    def insert_image(self, name, path):
        """커서 자리에 그림을 넣는다."""
        self.files[name] = path
        if self._put_image("insert", name):
            self.text.focus_set()
            self._changed()
            return True
        return False

    def insert_file(self, name, path):
        """커서 자리에 공통 첨부파일 표시를 넣는다."""
        self.files[name] = path
        self._put_chip("insert", "{첨부:" + name + "}")
        self.text.focus_set()
        self._changed()

    def insert_row_file_spot(self):
        """커서 자리에 개별 첨부파일 자리 표시를 넣는다(한 번만)."""
        if ROW_FILE_TOKEN in runs_to_text(self.get_runs()):
            return False
        self._put_chip("insert", ROW_FILE_TOKEN)
        self.text.focus_set()
        self._changed()
        return True

    def used_images(self):
        """본문에 쓰인 그림·공통 첨부 이름."""
        t = runs_to_text(self.get_runs())
        return [m.group(1).strip() for m in IMG_TOKEN_RE.finditer(t)] + [m.group(1).strip() for m in FILE_TOKEN_RE.finditer(t)]

    def _selection(self):
        try:
            return self.text.index("sel.first"), self.text.index("sel.last")
        except tk.TclError:
            return None, None

    def _each_char(self, start, end):
        idx = start
        while self.text.compare(idx, "<", end):
            nxt = self.text.index(f"{idx}+1c")
            yield idx, nxt
            idx = nxt

    def apply(self, attr, value):
        start, end = self._selection()
        if not start:
            return
        for a, b in self._each_char(start, end):
            st = self._style_at(a)
            st[attr] = value
            self._set_style(a, b, st)
        self._changed()

    def toggle(self, attr):
        start, end = self._selection()
        if not start:
            return
        chars = list(self._each_char(start, end))
        all_on = all(self._style_at(a)[attr] for a, _ in chars if self.text.get(a) != "\n")
        for a, b in chars:
            st = self._style_at(a)
            st[attr] = not all_on
            self._set_style(a, b, st)
        self._changed()

    def insert_text(self, text):
        idx = self.text.index("insert")
        prev = self.text.index(f"{idx}-1c") if self.text.compare(idx, ">", "1.0") else idx
        st = self._style_at(prev)
        self.text.insert("insert", text, (self._tag_for(st),))
        self.text.focus_set()
        self._changed()

    def _after_edit(self, event=None):
        # 새로 친 글자(서식 없음)는 바로 앞 글자의 서식을 따른다
        prev_st = None
        for a, b in self._each_char("1.0", "end-1c"):
            has = any(t.startswith("st_") for t in self.text.tag_names(a))
            if has:
                prev_st = self._style_at(a)
            else:
                self.text.tag_add(self._tag_for(prev_st or self._style_at(b)), a, b)
        self._changed()

    def _changed(self):
        if self.on_change:
            self.on_change()


def save_result_xlsx(path, sheet_title, cols, rows):
    """표 내용을 엑셀(xlsx)로 저장. 제목 줄 굵게·고정, 성공/실패 색 표시."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_title[:31]
    ws.append([c["header"] for c in cols])
    head_fill = PatternFill("solid", fgColor="DCE6F1")
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = head_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ok_fill, ng_fill = PatternFill("solid", fgColor="E2EFDA"), PatternFill("solid", fgColor="FCE4D6")
    status_idx = next((i for i, c in enumerate(cols) if c["key"] == "status"), None)
    for row in rows:
        ws.append([safe_text(row.get(c["key"], "")) for c in cols])
        if status_idx is not None:
            st = normalize_result_status(row.get("status", ""))
            cell = ws.cell(row=ws.max_row, column=status_idx + 1)
            if st == "성공":
                cell.fill = ok_fill
            elif st in ("실패", "중단", "확인필요"):
                cell.fill = ng_fill
    for i, c in enumerate(cols, start=1):
        longest = max([len(c["header"])] + [len(safe_text(r.get(c["key"], ""))) for r in rows[:500]])
        ws.column_dimensions[get_column_letter(i)].width = max(8, min(60, longest * 1.6 + 2))
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)


class MessengerBatchHelperUIv3(MessengerBatchHelperUI):

    def __init__(self, root):
        self.ui_settings = self._load_ui_settings()
        self.tpl = load_note_template()
        self.field_cols = list(self.tpl.get("columns") or DEFAULT_FIELD_COLUMNS)
        self.sent_log = NoteSentLog()
        self.preview_win = None
        super().__init__(root)
        self.root.title(f"메신저 작업 도우미 {APP_VERSION}")
        self.root.geometry("1400x960")
        self.root.minsize(1220, 820)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        try:
            self._icon_img = tk.PhotoImage(data=APP_ICON_PNG_B64)
            self.root.iconphoto(True, self._icon_img)
        except Exception:
            pass
        self._start_f12_watch()

    # ---------------- F12 중지 (어느 창이 앞에 있어도) ----------------
    def _start_f12_watch(self):
        """도우미 창이 뒤에 있어도 F12를 누르면 멈추도록 키 상태를 직접 살핀다."""
        def watch():
            try:
                get_state = ctypes.windll.user32.GetAsyncKeyState
            except Exception:
                return
            was_down = False
            while True:
                try:
                    down = bool(get_state(0x7B) & 0x8000)   # VK_F12
                    if down and not was_down and self.running:
                        self.stop_event.set()                # 작업 쪽에 바로 알림
                        self.event_q.put(("f12_stop", None))
                    was_down = down
                except Exception:
                    pass
                time.sleep(0.05)
        threading.Thread(target=watch, daemon=True).start()

    # ---------------- 설정 ----------------
    def _load_ui_settings(self):
        try:
            with open(NOTE_UI_SETTINGS_FILE, "r", encoding="utf-8") as f:
                return json.load(f) or {}
        except Exception:
            return {}

    def _save_ui_settings(self):
        try:
            s = {"test_name": self.test_name_var.get().strip(),
                 "test_dept": self.test_dept_var.get().strip(),
                 "skip_sent": bool(self.skip_sent_var.get())}
            with open(NOTE_UI_SETTINGS_FILE, "w", encoding="utf-8") as f:
                json.dump(s, f, ensure_ascii=False, indent=1)
        except Exception:
            pass

    def _store_template(self):
        """쪽지 내용(서식·그림·개별문구 열 이름)을 쪽지내용.json에 저장."""
        try:
            runs = self.editor.get_runs()
            used = set(self.editor.used_images())
            files = {k: v for k, v in self.editor.files.items() if k in used}
            self.tpl = {"runs": runs, "files": files, "columns": list(self.field_cols)}
            save_note_template(self.tpl)
        except Exception:
            pass

    def _on_close(self):
        if self.running and not messagebox.askyesno("종료 확인", "작업이 실행 중입니다. 그래도 닫을까요?"):
            return
        self._store_template()
        self._save_ui_settings()
        self.root.destroy()

    # ---------------- 쪽지 발송 탭 (2.0 화면 기준) ----------------
    def _build_note_tab(self):
        container = ttk.Frame(self.note_tab, padding=10)
        container.pack(fill="both", expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(2, weight=1)

        top_area = ttk.Frame(container)
        top_area.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        top_area.columnconfigure(0, weight=0, minsize=430)
        top_area.columnconfigure(1, weight=1)

        # 왼쪽: 안내 · 발송 설정 · 시험 발송
        guide = self._card_grid(top_area, 0, 0, sticky="nsew", padx=(0, 8))
        guide.columnconfigure(1, weight=1)
        ttk.Label(guide, text="쪽지 발송", style="CardTitle.TLabel").grid(row=0, column=0, columnspan=4, sticky="w")
        ttk.Label(guide, justify="left", text=(
            "- 표의 각 칸에 직접 입력하거나 엑셀·한글 표에서 복사해 Ctrl+V로 붙여넣습니다.\n"
            "- 검색어1·2로 찾고, 여러 명이면 검색어3으로 한 번 더 좁힙니다.\n"
            "- 개별문구 열: 표 제목을 누르면 이름 바꾸기, 우클릭하면 열 추가·삭제.\n"
            "  본문에 {{열 이름}}을 넣으면 사람마다 그 칸 내용으로 바뀝니다.\n"
            "- [그림 넣기]: 본문 커서 자리에 그림(QR 등)\n"
            "- [파일 넣기]: 모두에게 같은 파일을 커서 자리에 첨부\n"
            "- [개별 첨부 자리]: 표의 첨부파일 칸 파일이 들어갈 자리(없으면 맨 끝)")).grid(row=1, column=0, columnspan=4, sticky="w", pady=(6, 6))
        self.skip_sent_var = tk.BooleanVar(value=self.ui_settings.get("skip_sent", True))
        ttk.Checkbutton(guide, text="이미 보낸 쪽지는 건너뛰기", variable=self.skip_sent_var).grid(row=2, column=0, columnspan=2, sticky="w")
        ttk.Button(guide, text="발송 기록 지우기", command=self.clear_sent_log).grid(row=2, column=2, columnspan=2, sticky="e")
        ttk.Separator(guide).grid(row=3, column=0, columnspan=4, sticky="ew", pady=8)
        ttk.Label(guide, text="시험 발송 (표 위쪽 내용을 나에게 먼저)", style="CardTitle.TLabel").grid(row=4, column=0, columnspan=4, sticky="w")
        ttk.Label(guide, text="받을 사람").grid(row=5, column=0, sticky="w", pady=2)
        self.test_name_var = tk.StringVar(value=self.ui_settings.get("test_name", ""))
        ttk.Entry(guide, textvariable=self.test_name_var, width=14).grid(row=5, column=1, sticky="w")
        ttk.Label(guide, text="부서").grid(row=5, column=2, sticky="e", padx=(6, 4))
        self.test_dept_var = tk.StringVar(value=self.ui_settings.get("test_dept", ""))
        ttk.Entry(guide, textvariable=self.test_dept_var, width=14).grid(row=5, column=3, sticky="w")
        ttk.Label(guide, text="건수").grid(row=6, column=0, sticky="w", pady=2)
        self.test_count_var = tk.IntVar(value=1)
        ttk.Spinbox(guide, from_=1, to=5, textvariable=self.test_count_var, width=5).grid(row=6, column=1, sticky="w")
        ttk.Button(guide, text="시험 발송", command=self.start_note_test).grid(row=6, column=3, sticky="e")

        # 오른쪽: 공통 쪽지 내용
        msg_card = self._card_grid(top_area, 0, 1, sticky="nsew", padx=(8, 0))
        msg_card.columnconfigure(0, weight=1)
        head = ttk.Frame(msg_card, style="Card.TFrame")
        head.grid(row=0, column=0, sticky="ew")
        ttk.Label(head, text="공통 쪽지 내용", style="CardTitle.TLabel").pack(side="left")
        ttk.Button(head, text="개별 첨부 자리", width=12, command=self.insert_row_file_spot).pack(side="right", padx=1)
        ttk.Button(head, text="파일 넣기", width=9, command=self.insert_body_file).pack(side="right", padx=1)
        ttk.Button(head, text="그림 넣기", width=9, command=self.insert_body_image).pack(side="right", padx=1)
        self.chip_frame = tk.Frame(msg_card, bg="#ffffff")
        self.chip_frame.grid(row=1, column=0, sticky="ew", pady=(6, 2))
        self.editor = RichTemplateEditor(msg_card, height=11, on_change=self._on_template_change)
        self.editor.grid(row=2, column=0, sticky="nsew")
        self.tpl_warn_var = tk.StringVar(value="")
        ttk.Label(msg_card, textvariable=self.tpl_warn_var, foreground="#b91c1c", wraplength=800).grid(row=3, column=0, sticky="w", pady=(4, 0))

        # 상태 줄
        info = self._card_grid(container, 1, sticky="ew", pady=(0, 8))
        info.columnconfigure(5, weight=1)
        self.note_status_var = tk.StringVar(value="상태: 대기")
        self.note_current_var = tk.StringVar(value="현재: -")
        self.note_summary_var = tk.StringVar(value="성공 0 / 실패 0 / 중단 0")
        ttk.Label(info, textvariable=self.note_status_var).grid(row=0, column=0, sticky="w", columnspan=2)
        ttk.Label(info, textvariable=self.note_current_var, wraplength=620).grid(row=0, column=2, sticky="w", columnspan=2)
        ttk.Label(info, textvariable=self.note_summary_var).grid(row=0, column=4, sticky="e", padx=(10, 0))

        # 수신자 표
        self.note_card = self._card_grid(container, 2, sticky="nsew")
        self.note_card.columnconfigure(0, weight=1)
        self.note_card.rowconfigure(1, weight=1)
        ttk.Label(self.note_card, text="수신자 표", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        bar = ttk.Frame(self.note_card)
        bar.grid(row=0, column=1, sticky="e")
        self.note_card.columnconfigure(1, weight=0)
        ttk.Button(bar, text="초기화", command=self.reset_note_table).pack(side="left", padx=3)
        ttk.Button(bar, text="행 추가", command=lambda: self.note_grid.add_row()).pack(side="left", padx=3)
        ttk.Button(bar, text="선택 행 삭제", command=self.delete_note_selected).pack(side="left", padx=3)
        ttk.Button(bar, text="첨부파일 선택", command=self.select_attachment_for_note_row).pack(side="left", padx=2)
        ttk.Button(bar, text="미리보기", command=self.preview_note).pack(side="left", padx=3)
        ttk.Button(bar, text="검색만 확인", command=self.start_note_check).pack(side="left", padx=3)
        ttk.Button(bar, text="실패만 남기기", command=self.keep_note_failures).pack(side="left", padx=2)
        ttk.Button(bar, text="결과 저장", command=self.export_note_result).pack(side="left", padx=3)
        ttk.Button(bar, text="실제 발송 실행", style="Accent.TButton", command=self.start_note_send).pack(side="left", padx=(16, 3))
        ttk.Button(bar, text="중지", command=self.stop).pack(side="left", padx=3)
        self.note_grid = None
        self._make_note_grid()

        self.editor.set_runs(self.tpl.get("runs", []), self.tpl.get("files", {}))
        self._refresh_chips()

    # ---------------- 표 ----------------
    def _note_columns(self):
        cols = [{"key": "k1", "header": "검색어1", "width": 12, "editable": True, "zone": "search"},
                {"key": "k2", "header": "검색어2", "width": 16, "editable": True, "zone": "search"},
                {"key": "k3", "header": "검색어3", "width": 10, "editable": True, "zone": "search"}]
        for i, name in enumerate(self.field_cols):
            cols.append({"key": f"f{i}", "header": name, "width": max(10, min(24, len(name) * 2 + 4)),
                         "editable": True, "zone": "field", "kind": "field"})
        cols += [{"key": "attach", "header": "첨부파일", "width": 18, "editable": True, "zone": "field", "kind": "attach"},
                 {"key": "status", "header": "성공 여부", "width": 9, "editable": False, "zone": "result"},
                 {"key": "reason", "header": "실패 사유 / 처리 결과", "width": 30, "editable": False, "zone": "result"}]
        return cols

    def _make_note_grid(self, rows=None):
        if self.note_grid is not None:
            self.note_grid.destroy()
        self.note_grid = ZonedGrid(self.note_card, self._note_columns(), rows=12, on_paste=self._handle_note_paste,
                                   on_add=self._field_add, on_delete=self._field_delete, on_rename=self._field_rename)
        self.note_grid.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(8, 0))
        if rows:
            self.note_grid.set_rows(rows)
        g = self.note_grid
        g.after(150, lambda: g.winfo_exists() and g._xview("moveto", 0))

    def _row_values(self, row):
        """본문 자리표시에 쓸 값: 개별문구 열(열 이름 그대로). 검색어는 메신저 검색에만 씀."""
        v = {}
        for i, name in enumerate(self.field_cols):
            v[name] = safe_text(row.get(f"f{i}", ""))
        return v

    def _row_empty(self, row):
        keys = ["k1", "k2", "k3", "attach"] + [f"f{i}" for i in range(len(self.field_cols))]
        return not any(safe_text(row.get(k, "")).strip() for k in keys)

    def _handle_note_paste(self, grid, start_row, start_key, text):
        if self.running:
            messagebox.showwarning("안내", "실행 중에는 붙여넣을 수 없습니다.")
            return
        matrix = get_clipboard_excel_matrix(text)
        if not matrix:
            matrix = [[safe_text(text)]]
        grid.paste_matrix(start_row, start_key, matrix)
        self._update_note_summary()

    def reset_note_table(self):
        if self.running:
            return
        if messagebox.askyesno("초기화 확인", "쪽지 발송 표의 모든 입력값과 처리 결과를 초기화할까요?"):
            self._make_note_grid()
            self.note_status_var.set("상태: 대기")
            self.note_current_var.set("현재: -")
            self._update_note_summary()

    def delete_note_selected(self):
        if self.running:
            return
        self.note_grid.delete_selected()
        if not self.note_grid.data:
            self.note_grid.add_row()
        self._update_note_summary()

    def select_attachment_for_note_row(self):
        if self.running:
            return
        row = self.note_grid.selected_row
        if row is None:
            messagebox.showinfo("안내", "첨부파일을 넣을 행을 먼저 고르세요.")
            return
        files = filedialog.askopenfilenames(title="이 행에 첨부할 파일")
        if not files:
            return
        self.note_grid.set_value(row, "attach", ";".join(_windows_style_path(p) for p in files))
        self.note_grid.clear_result(row)

    def keep_note_failures(self):
        if self.running:
            return
        keeps = [dict(r) for r in self.note_grid.data
                 if not self._row_empty(r) and normalize_result_status(r.get("status")) != "성공"]
        self.note_grid.set_rows(keeps)
        self._update_note_summary()
        messagebox.showinfo("안내", f"성공이 아닌 {len(keeps)}건만 남겼습니다.")

    def export_note_result(self):
        rows = [r for r in self.note_grid.data if not self._row_empty(r)]
        if not rows:
            messagebox.showwarning("결과 저장", "저장할 내용이 없습니다.")
            return
        path = filedialog.asksaveasfilename(title="결과 저장", defaultextension=".xlsx",
                                            initialfile=f"쪽지발송결과_{datetime.now():%Y%m%d_%H%M}.xlsx",
                                            filetypes=[("엑셀", "*.xlsx")])
        if not path:
            return
        try:
            save_result_xlsx(path, "쪽지발송결과", self.note_grid.columns_def, rows)
            messagebox.showinfo("결과 저장", path)
        except Exception as e:
            messagebox.showerror("결과 저장 실패", str(e))

    # ---------------- 개별문구 열 ----------------
    def _field_items(self):
        return [[f"f{i}", n] for i, n in enumerate(self.field_cols)]

    def _field_name_ok(self, name, skip_key=None):
        n = name.strip()
        if not n or any(ch in n for ch in "{}\n") or n.startswith("그림:"):
            messagebox.showwarning("열 이름", "열 이름을 입력하세요. ({ } 는 쓸 수 없습니다)")
            return False
        taken = {"검색어1", "검색어2", "검색어3", "첨부파일", "성공 여부", "개별첨부"} | {
            nm for k, nm in self._field_items() if k != skip_key}
        if n in taken:
            messagebox.showwarning("열 이름", f"'{n}'은(는) 이미 있는 열 이름입니다.")
            return False
        return True

    def _field_add(self, key, side):
        if self.running:
            return
        items = self._field_items()
        k = len(items) + 1
        while any(n == f"개별문구{k}" for _, n in items):
            k += 1
        if key == "attach":
            pos = len(items)
        else:
            pos = next(i for i, (kk, _) in enumerate(items) if kk == key) + (1 if side == "right" else 0)
        items.insert(pos, [None, f"개별문구{k}"])
        self._apply_field_columns(items)
        self.note_grid.after(80, lambda: self.note_grid.start_header_rename(f"f{pos}"))

    def _field_delete(self, key):
        if self.running:
            return
        items = self._field_items()
        if len(items) <= 1:
            return
        name = dict(items)[key]
        has = any(safe_text(r.get(key, "")).strip() for r in self.note_grid.data)
        if has and not messagebox.askyesno("열 삭제", f"'{name}' 열에 입력한 내용도 함께 지워집니다. 삭제할까요?"):
            return
        self._apply_field_columns([it for it in items if it[0] != key])

    def _field_rename(self, key, name):
        if self.running or not self._field_name_ok(name, skip_key=key):
            self.note_grid._redraw()
            return
        items = self._field_items()
        for it in items:
            if it[0] == key:
                it[1] = name.strip()
        self._apply_field_columns(items)

    def _apply_field_columns(self, items):
        old_names = {f"f{i}": n for i, n in enumerate(self.field_cols)}
        renames = {old_names[k]: n for k, n in items if k and old_names.get(k) != n}
        removed = [n for k, n in old_names.items() if k not in {x[0] for x in items}]
        # 표 내용은 열을 따라 옮긴다
        rows = []
        for r in self.note_grid.data:
            nr = {k: r.get(k, "") for k in ("k1", "k2", "k3", "attach", "status", "reason")}
            for j, (k, _) in enumerate(items):
                nr[f"f{j}"] = r.get(k, "") if k else ""
            rows.append(nr)
        self.field_cols = [n for _, n in items]
        # 본문 자리표시 이름도 같이 바꾼다
        if renames:
            runs = []
            for text, st in self.editor.get_runs():
                text = PLACEHOLDER_RE.sub(lambda m: ph(renames.get(m.group(1).strip(), m.group(1))), text)
                runs.append([text, st])
            self.editor.set_runs(runs)
        self._make_note_grid(rows)
        self._refresh_chips()
        self._store_template()
        used = set(template_placeholders(self.editor.get_runs()))
        gone = [n for n in removed if n in used]
        if gone:
            messagebox.showwarning("개별문구 열", "지운 열이 본문에 아직 있습니다: " + ", ".join(ph(n) for n in gone))

    def _refresh_chips(self):
        for w in self.chip_frame.winfo_children():
            w.destroy()
        tk.Label(self.chip_frame, text="자리표시:", bg="#ffffff", fg="#52616f", font=("Malgun Gothic", 9)).pack(side="left")
        for h in list(self.field_cols):
            tk.Button(self.chip_frame, text=ph(h), font=("Malgun Gothic", 9), fg="#1d4ed8", bg="#eff6ff",
                      relief="flat", cursor="hand2", command=lambda x=h: self.editor.insert_text(ph(x))).pack(side="left", padx=2)
        self._on_template_change()

    # ---------------- 본문 그림 ----------------
    def insert_body_image(self):
        path = filedialog.askopenfilename(title="본문에 넣을 그림", filetypes=[("그림", "*.png *.jpg *.jpeg *.gif *.bmp"), ("모든 파일", "*.*")])
        if not path:
            return
        base = os.path.basename(path)
        name, n = base, 2
        while name in self.editor.files and os.path.normcase(self.editor.files[name]) != os.path.normcase(path):
            stem, ext = os.path.splitext(base)
            name = f"{stem}({n}){ext}"
            n += 1
        if not self.editor.insert_image(name, _windows_style_path(path)):
            messagebox.showwarning("그림 넣기", "그림을 열지 못했습니다. png·jpg·gif·bmp 파일인지 확인하세요.")

    def insert_body_file(self):
        paths = filedialog.askopenfilenames(title="모두에게 첨부할 파일(커서 자리에 들어감)")
        for path in paths or []:
            base = os.path.basename(path)
            name, n = base, 2
            while name in self.editor.files and os.path.normcase(self.editor.files[name]) != os.path.normcase(_windows_style_path(path)):
                stem, ext = os.path.splitext(base)
                name = f"{stem}({n}){ext}"
                n += 1
            self.editor.insert_file(name, _windows_style_path(path))

    def insert_row_file_spot(self):
        if not self.editor.insert_row_file_spot():
            messagebox.showinfo("개별 첨부 자리", "개별 첨부 자리는 본문에 이미 있습니다. 옮기려면 지우고 다시 넣으세요.")

    def clear_sent_log(self):
        if not messagebox.askyesno("발송 기록 지우기", "쪽지 발송 기록을 모두 지웁니다.\n지우면 같은 쪽지를 다시 보낼 수 있습니다. 계속할까요?"):
            return
        n = self.sent_log.clear_source("쪽지")
        messagebox.showinfo("발송 기록", f"{n}건의 기록을 지웠습니다.")

    def _on_template_change(self):
        if not hasattr(self, "tpl_warn_var"):
            return
        self.tpl_warn_var.set(" · ".join(self._template_problems()))

    def _template_problems(self):
        probs = []
        runs = self.editor.get_runs()
        if not runs_to_text(runs).strip():
            probs.append("본문이 비어 있음")
        known = set(self.field_cols)
        missing = sorted(set(k for k in template_placeholders(runs) if k not in known))
        if missing:
            probs.append("표에 없는 자리표시: " + ", ".join(ph(k) for k in missing))
        for name in self.editor.used_images():
            p = self.editor.files.get(name)
            if not p or not os.path.exists(p):
                probs.append(f"파일 없음: {name}")
        return probs

    # ---------------- 발송 준비 ----------------
    def _send_key(self, kws, parts, attach):
        sig = json.dumps([runs_to_text(v) if k == "text" else v for k, v in parts] + [attach] + list(kws), ensure_ascii=False)
        h = hashlib.md5(sig.encode("utf-8")).hexdigest()[:12]
        return NoteSentLog.key("쪽지", h, kws[0], kws[1])

    def _collect_note_items(self, mark=True):
        runs_tpl = self.editor.get_runs()
        files = dict(self.editor.files)
        items, invalid, skipped = [], 0, 0
        for idx, row in enumerate(self.note_grid.data):
            if self._row_empty(row):
                continue
            values = self._row_values(row)
            kws = [safe_text(row.get(k, "")).strip() for k in ("k1", "k2", "k3")]
            if not kws[0]:
                if mark:
                    self.note_grid.set_result(idx, "실패", "검색어1 없음")
                invalid += 1
                continue
            attach = safe_text(row.get("attach", "")).strip()
            if attach:
                _, miss = validate_attachment_paths(attach)
                if miss:
                    if mark:
                        self.note_grid.set_result(idx, "실패", "첨부파일 없음: " + os.path.basename(miss[0]))
                    invalid += 1
                    continue
            runs = fill_runs(runs_tpl, values)
            row_files = split_attachment_paths(attach)
            parts, common = [], set()
            body_parts = split_body_parts(runs)
            for kind, val in body_parts:
                if kind in ("image", "file"):
                    path = files.get(val, "")
                    parts.append((kind, path))
                    if kind == "file":
                        common.add(os.path.normcase(path))
                elif kind == "rowfile":
                    parts.append(("rowfile", None))
                else:
                    parts.append((kind, val))
            row_files = [p for p in row_files if os.path.normcase(p) not in common]   # 공통과 같은 파일은 한 번만
            if not any(k == "rowfile" for k, _ in parts):
                parts.append(("rowfile", None))                                  # 자리가 없으면 맨 끝
            expanded = []
            for kind, val in parts:
                if kind == "rowfile":
                    expanded += [("file", p) for p in row_files]
                else:
                    expanded.append((kind, val))
            parts = expanded
            key = self._send_key(kws, parts, attach)
            if self.skip_sent_var.get() and self.sent_log.get(key):
                skipped += 1
                if mark:
                    info = self.sent_log.get(key)
                    self.note_grid.set_result(idx, "성공", f"이전에 보냄 ({info.get('t', '')}) - 건너뜀")
                continue
            items.append({"row": idx, "keywords": kws, "label": " / ".join(k for k in kws if k),
                          "runs": runs, "parts": parts, "attach": attach, "key": key})
        return items, invalid, skipped

    def _check_before_send(self):
        self._store_template()
        probs = self._template_problems()
        if probs:
            messagebox.showwarning("내용 확인", "아래 문제를 먼저 고쳐 주세요.\n\n- " + "\n- ".join(probs))
            return False
        return True

    def preview_note(self):
        items, _, _ = self._collect_note_items(mark=False)
        if not items:
            messagebox.showwarning("안내", "미리 볼 대상이 없습니다.")
            return
        start = 0
        sel = self.note_grid.selected_row or 0
        for i, it in enumerate(items):
            if it["row"] >= sel:
                start = i
                break
        if self.preview_win is not None and self.preview_win.winfo_exists():
            self.preview_win.destroy()
        win = tk.Toplevel(self.root)
        self.preview_win = win
        win.title("쪽지 미리보기")
        win.geometry("780x760")
        win.columnconfigure(0, weight=1)
        win.rowconfigure(2, weight=1)
        pos_var, meta_var = tk.StringVar(), tk.StringVar()
        nav = ttk.Frame(win, padding=8)
        nav.grid(row=0, column=0, sticky="ew")
        state = {"i": start}
        view = RichTemplateEditor(win, height=26, readonly=True, toolbar=False)
        files = dict(self.editor.files)

        def show():
            it = items[state["i"]]
            pos_var.set(f"{state['i'] + 1} / {len(items)}  (표 {it['row'] + 1}행)")
            att = [os.path.basename(v) for k, v in it["parts"] if k == "file"]
            meta_var.set(f"검색어: {it['label']}\n첨부파일(순서대로): {', '.join(att) if att else '없음'}")
            view.set_runs(it["runs"], files)

        ttk.Button(nav, text="◀", width=4, command=lambda: (state.__setitem__("i", (state["i"] - 1) % len(items)), show())).pack(side="left")
        ttk.Label(nav, textvariable=pos_var, width=26, anchor="center").pack(side="left")
        ttk.Button(nav, text="▶", width=4, command=lambda: (state.__setitem__("i", (state["i"] + 1) % len(items)), show())).pack(side="left")
        ttk.Label(win, textvariable=meta_var, padding=(10, 0)).grid(row=1, column=0, sticky="w")
        view.grid(row=2, column=0, sticky="nsew", padx=10, pady=10)
        show()

    def start_note_send(self):
        if self.running:
            messagebox.showwarning("안내", "이미 실행 중입니다.")
            return
        if not self._check_before_send():
            return
        items, invalid, skipped = self._collect_note_items()
        self.note_grid._redraw()
        if not items:
            messagebox.showwarning("입력 확인", f"보낼 대상이 없습니다.\n(이미 보냄 {skipped}건, 입력 오류 {invalid}건)")
            return
        first = items[0]
        msg = f"총 {len(items)}건의 쪽지를 실제로 보냅니다.\n\n첫 번째: 검색어 {first['label']}"
        if skipped:
            msg += f"\n이미 보낸 {skipped}건은 건너뜁니다."
        if invalid:
            msg += f"\n입력 오류 {invalid}건은 보내지 않습니다(표의 실패 사유 확인)."
        msg += "\n검색 결과가 없거나 여러 명인 사람에게는 보내지 않습니다.\n\n계속할까요?"
        if not messagebox.askyesno("실제 발송 확인", msg):
            return
        self._run_note_worker(items, mode="send")

    def start_note_test(self):
        if self.running:
            return
        name = self.test_name_var.get().strip()
        dept = self.test_dept_var.get().strip()
        if not name:
            messagebox.showwarning("시험 발송", "시험으로 받을 사람 이름(보통 본인)을 입력하세요.")
            return
        if not self._check_before_send():
            return
        self._save_ui_settings()
        items, _, _ = self._collect_note_items(mark=False)
        if not items:
            messagebox.showwarning("시험 발송", "표에 내용이 없습니다.")
            return
        n = max(1, min(5, int(self.test_count_var.get() or 1)))
        tests = [dict(it, keywords=[name, dept, ""], label=f"{name} / {dept}", row=None, key=None) for it in items[:n]]
        if not messagebox.askyesno("시험 발송", f"표 위쪽 {len(tests)}건의 내용을 '{name} / {dept or '부서 없음'}'에게 보냅니다.\n계속할까요?"):
            return
        self._run_note_worker(tests, mode="test")

    def start_note_check(self):
        if self.running:
            return
        items, invalid, skipped = self._collect_note_items()
        self.note_grid._redraw()
        if not items:
            messagebox.showwarning("검색만 확인", "확인할 대상이 없습니다.")
            return
        mins = max(1, round(len(items) * 3 / 60))
        if not messagebox.askyesno("검색만 확인", f"{len(items)}건을 메신저에서 검색만 해 봅니다. 쪽지는 보내지 않습니다.\n(약 {mins}분, 실행 중에는 마우스·키보드 사용 금지)\n\n계속할까요?"):
            return
        self._run_note_worker(items, mode="check")

    def _run_note_worker(self, items, mode):
        self.stop_event.clear()
        self.running = True
        self.note_status_var.set({"test": "상태: 시험 발송 중", "check": "상태: 검색 확인 중"}.get(mode, "상태: 발송 중"))
        self.note_current_var.set("현재: 준비 중")
        self._save_ui_settings()
        self.worker = threading.Thread(target=self._note_worker, args=(items, mode), daemon=True)
        self.worker.start()

    def _note_worker(self, items, mode="send"):
        logf = open_run_log("쪽지_" + mode)
        with redirect_stdout(QueueWriter(file_obj=logf)):
            counts = {"성공": 0, "실패": 0, "확인필요": 0}
            try:
                main_hwnd = find_messenger_hwnd()
                if not main_hwnd:
                    self.event_q.put(("fatal", "JBEdu Messenger+ 창을 찾지 못했습니다."))
                    return
                sess = MessengerSession(main_hwnd, self.stop_event)
                if not sess.bring_front():
                    self.event_q.put(("fatal", "메신저 창을 앞으로 가져오지 못했습니다.\n메신저를 한 번 눌러 앞에 둔 뒤 다시 실행하세요."))
                    return
                total = len(items)
                stopped = False
                for idx, item in enumerate(items, start=1):
                    row = item["row"]
                    self.event_q.put(("note_current", {"row": row, "text": f"현재: {idx}/{total} - {item.get('label', '')}"}))
                    print(f"\n######## {idx}/{total} {item.get('label', '')} ########")
                    try:
                        person, judg = v3_find_person(sess, item["keywords"])
                        if person is None:
                            status, reason = "실패", judg.get("reason", "검색 실패")
                        elif mode == "check":
                            status, reason = "확인됨", judg.get("person", "")
                        else:
                            status, reason = v3_send_note_parts(sess, person, item["parts"], actually_send=True)
                            reason = f"{judg.get('person', '')} · {reason}"
                    except ForegroundLost as e:
                        print(f"앞창 잃음: {e}")
                        if row is not None:
                            self.event_q.put(("note_row", {"row": row, "status": "중단", "reason": "메신저가 앞에 없어서 멈춤"}))
                        self.event_q.put(("paused", "메신저 창이 앞에 없어서 멈췄습니다.\n\n실행 중에는 마우스·키보드를 쓰지 마세요.\n다시 실행하면 보낸 쪽지는 건너뛰고 이어서 합니다."))
                        return
                    except StopRequested:
                        if row is not None:
                            self.event_q.put(("note_row", {"row": row, "status": "중단", "reason": "사용자 중지"}))
                        stopped = True
                        break
                    except Exception as e:
                        traceback.print_exc()
                        status, reason = "실패", f"예외 발생: {e.__class__.__name__}"
                    print(f"결과: {status} - {reason}")
                    if mode == "send" and item.get("key") and status in ("성공", "확인필요"):
                        self.sent_log.mark(item["key"], {"t": datetime.now().strftime("%Y-%m-%d %H:%M"),
                                                         "name": item.get("label", ""), "unclear": status == "확인필요"})
                    key = "성공" if status in ("성공", "확인됨") else ("확인필요" if status == "확인필요" else "실패")
                    counts[key] += 1
                    if row is not None:
                        self.event_q.put(("note_row", {"row": row, "status": status, "reason": reason}))
                    else:
                        self.event_q.put(("note_test_row", f"{idx}/{total} {status} - {reason}"))
                tail = " (중지됨)" if stopped else ""
                if mode == "test":
                    msg = f"시험 발송 끝{tail}: 성공 {counts['성공']}, 실패 {counts['실패']}, 확인필요 {counts['확인필요']}\n받은 쪽지의 서식·그림·첨부를 확인하세요."
                elif mode == "check":
                    msg = f"검색 확인 끝{tail}: 1명 확인 {counts['성공']}, 실패 {counts['실패']}\n실패한 행은 검색어를 고친 뒤 다시 확인하세요."
                else:
                    msg = f"쪽지 발송 끝{tail}: 성공 {counts['성공']}, 실패 {counts['실패']}, 확인필요 {counts['확인필요']}"
                self.event_q.put(("note_done", msg))
            except Exception:
                traceback.print_exc()
                self.event_q.put(("fatal", "실행 중 예외가 발생했습니다. logs 폴더의 작업 기록을 확인하세요."))
            finally:
                try:
                    logf.close()
                except Exception:
                    pass

    def _refresh_note_info(self):
        self._update_note_summary()

    # ---------------- 사용자그룹 탭 ----------------
    def _build_group_tab(self):
        container = ttk.Frame(self.group_tab, padding=10)
        container.pack(fill="both", expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(2, weight=1)

        top = self._card_grid(container, 0, sticky="ew", pady=(0, 8))
        top.columnconfigure(3, weight=1)
        ttk.Label(top, text="사용자그룹 추가", style="CardTitle.TLabel").grid(row=0, column=0, columnspan=4, sticky="w")
        ttk.Label(top, text="- 엑셀을 열거나 표에 붙여넣기. 검색어1·2로 찾고, 여러 명이면 검색어3으로 한 번 더 좁힘 (검색 필터는 '전체')\n"
                            f"- '{WORK_GROUP_NAME}' 그룹을 만들어 넣음. 같은 이름의 그룹이 이미 있으면 멈춤 → 끝난 뒤 메신저에서 그룹 이름을 바꿔 두세요\n"
                            "- 실행 중 메신저가 앞에서 사라지면 자동으로 멈춤. 다시 실행하면 '성공' 행은 건너뜀",
                  justify="left").grid(row=1, column=0, columnspan=4, sticky="w", pady=(4, 6))
        ttk.Label(top, text="작업용 그룹이 이미 있으면").grid(row=3, column=0, sticky="w", pady=(6, 0))
        self.group_exist_var = tk.StringVar(value="stop")
        rf = ttk.Frame(top, style="Card.TFrame")
        rf.grid(row=3, column=1, columnspan=3, sticky="w", pady=(6, 0))
        ttk.Radiobutton(rf, text="멈춤", value="stop", variable=self.group_exist_var).pack(side="left")
        ttk.Radiobutton(rf, text="그 그룹에 이어서 추가 (중간에 멈췄던 작업을 계속할 때)", value="resume", variable=self.group_exist_var).pack(side="left", padx=14)

        info = self._card_grid(container, 1, sticky="ew", pady=(0, 8))
        info.columnconfigure(3, weight=1)
        self.group_status_var = tk.StringVar(value="상태: 대기")
        self.group_current_var = tk.StringVar(value="현재: -")
        self.group_summary_var = tk.StringVar(value="성공 0 / 실패 0 / 중단 0")
        ttk.Label(info, textvariable=self.group_status_var).grid(row=0, column=0, sticky="w")
        ttk.Label(info, textvariable=self.group_current_var, wraplength=640).grid(row=0, column=1, sticky="w", padx=(12, 0))
        ttk.Label(info, textvariable=self.group_summary_var).grid(row=0, column=4, sticky="e")

        card = self._card_grid(container, 2, sticky="nsew")
        card.columnconfigure(0, weight=1)
        card.rowconfigure(1, weight=1)
        ttk.Label(card, text="대상자 표", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        bar = ttk.Frame(card)
        bar.grid(row=0, column=1, sticky="e")
        ttk.Button(bar, text="엑셀/CSV 열기", command=self.open_group_file).pack(side="left", padx=3)
        ttk.Button(bar, text="초기화", command=self.reset_group_table).pack(side="left", padx=3)
        ttk.Button(bar, text="행 추가", command=lambda: self.group_grid.add_row()).pack(side="left", padx=3)
        ttk.Button(bar, text="선택 행 삭제", command=self.delete_group_selected).pack(side="left", padx=3)
        ttk.Button(bar, text="실패만 남기기", command=self.keep_group_failures).pack(side="left", padx=3)
        ttk.Button(bar, text="검색만 확인", command=self.start_group_check).pack(side="left", padx=3)
        ttk.Button(bar, text="결과 저장", command=self.export_group_result).pack(side="left", padx=3)
        ttk.Button(bar, text="전체 실행", style="Accent.TButton", command=self.start_group).pack(side="left", padx=(16, 3))
        ttk.Button(bar, text="중지", command=self.stop).pack(side="left", padx=3)
        self.group_grid = FastGrid(card, [
            {"key": "k1", "header": "검색어1", "width": 16, "editable": True},
            {"key": "k2", "header": "검색어2", "width": 24, "editable": True},
            {"key": "k3", "header": "검색어3 (여러 명일 때)", "width": 18, "editable": True},
            {"key": "status", "header": "성공 여부", "width": 9, "editable": False},
            {"key": "reason", "header": "찾은 사람 / 실패 사유", "width": 60, "editable": False},
        ], rows=12, on_paste=self._handle_group_paste)
        self.group_grid.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(8, 0))

    def open_group_file(self):
        if self.running:
            return
        path = filedialog.askopenfilename(title="대상자 명단 열기", filetypes=[("엑셀/CSV", "*.xlsx *.xlsm *.csv"), ("모든 파일", "*.*")])
        if not path:
            return
        try:
            headers, rows, sheets, sheet = load_table_file(path)
        except Exception as e:
            messagebox.showerror("명단 열기 실패", str(e))
            return
        pick = ColumnPickDialog(self.root, headers, rows).result
        if not pick:
            return
        kcols, fcol, fval = pick
        out = []
        for r in rows:
            c = r["cells"]
            if fcol and fval is not None and c.get(fcol, "") != fval:
                continue
            vals = [c.get(k, "") if k else "" for k in kcols]
            if any(vals):
                out.append({"k1": vals[0], "k2": vals[1], "k3": vals[2]})
        self.group_grid.set_rows(out)
        self.group_status_var.set(f"상태: {os.path.basename(path)}에서 {len(out)}명 불러옴")
        self._update_group_summary()

    def reset_group_table(self):
        if self.running:
            return
        if messagebox.askyesno("초기화 확인", "대상자 표의 입력값과 처리 결과를 모두 지울까요?"):
            self.group_grid.set_rows([{} for _ in range(12)])
            self.group_status_var.set("상태: 대기")
            self.group_current_var.set("현재: -")
            self._update_group_summary()

    def delete_group_selected(self):
        if self.running:
            return
        self.group_grid.delete_selected()
        self._update_group_summary()

    def keep_group_failures(self):
        if self.running:
            return
        keeps = [dict(r) for r in self.group_grid.data if normalize_result_status(r.get("status")) not in ("", "성공")]
        self.group_grid.set_rows(keeps)
        self._update_group_summary()
        messagebox.showinfo("안내", f"실패/중단 {len(keeps)}건만 남겼습니다.")

    def export_group_result(self):
        cols = self.group_grid.columns_def
        rows = [r for r in self.group_grid.data if any(safe_text(r.get(c["key"], "")) for c in cols)]
        if not rows:
            messagebox.showwarning("결과 저장", "저장할 내용이 없습니다.")
            return
        path = filedialog.asksaveasfilename(title="결과 저장", defaultextension=".xlsx",
                                            initialfile=f"사용자그룹결과_{datetime.now():%Y%m%d_%H%M}.xlsx",
                                            filetypes=[("엑셀", "*.xlsx")])
        if not path:
            return
        try:
            save_result_xlsx(path, "사용자그룹결과", cols, rows)
            messagebox.showinfo("결과 저장", path)
        except Exception as e:
            messagebox.showerror("결과 저장 실패", str(e))

    def _collect_group_items(self, skip_done=True):
        items, skipped = [], 0
        for idx, row in enumerate(self.group_grid.data):
            kws = [safe_text(row.get(k, "")).strip() for k in ("k1", "k2", "k3")]
            if not any(kws):
                continue
            st = safe_text(row.get("status", ""))
            if skip_done and st == "성공":
                skipped += 1
                continue
            if not kws[0]:
                self.group_grid.set_value(idx, "status", "실패", redraw=False)
                self.group_grid.set_value(idx, "reason", "검색어1 없음", redraw=False)
                continue
            items.append({"row": idx, "keywords": kws, "label": " / ".join(k for k in kws if k)})
        self.group_grid._redraw()
        return items, skipped

    def start_group_check(self):
        if self.running:
            return
        items, _ = self._collect_group_items(skip_done=False)
        if not items:
            messagebox.showwarning("검색만 확인", "확인할 대상이 없습니다.")
            return
        if not messagebox.askyesno("검색만 확인", f"{len(items)}명을 검색만 해 봅니다. 그룹에는 넣지 않습니다.\n계속할까요?"):
            return
        self._run_group_worker(items, "check", "")

    def start_group(self):
        if self.running:
            messagebox.showwarning("안내", "이미 실행 중입니다.")
            return
        items, skipped = self._collect_group_items()
        if not items:
            messagebox.showwarning("입력 확인", f"실행할 대상이 없습니다. (이미 성공 {skipped}명)")
            return
        msg = f"{len(items)}명을 '{WORK_GROUP_NAME}' 그룹에 추가합니다."
        if skipped:
            msg += f"\n이미 성공한 {skipped}명은 건너뜁니다."
        if self.group_exist_var.get() == "resume":
            msg += "\n같은 이름의 그룹이 있으면 그 그룹에 이어서 넣습니다."
        else:
            msg += "\n같은 이름의 그룹이 이미 있으면 멈춥니다."
        msg += "\n\n실행 중에는 마우스·키보드를 쓰지 마세요. 계속할까요?"
        if not messagebox.askyesno("실행 확인", msg):
            return
        self._run_group_worker(items, "add")

    def _run_group_worker(self, items, mode, final=""):
        opts = {"exist": self.group_exist_var.get()}
        self.stop_event.clear()
        self.running = True
        self.group_status_var.set("상태: 검색 확인 중" if mode == "check" else "상태: 실행 중")
        self.group_current_var.set("현재: 준비 중")
        self.worker = threading.Thread(target=self._group_worker, args=(items, mode, opts), daemon=True)
        self.worker.start()

    def _group_worker(self, items, mode="add", opts=None):
        opts = opts or {}
        logf = open_run_log("그룹_" + mode)
        with redirect_stdout(QueueWriter(file_obj=logf)):
            counts = {"성공": 0, "실패": 0}
            try:
                main_hwnd = find_messenger_hwnd()
                if not main_hwnd:
                    self.event_q.put(("fatal", "JBEdu Messenger+ 창을 찾지 못했습니다."))
                    return
                sess = MessengerSession(main_hwnd, self.stop_event)
                if not sess.bring_front():
                    self.event_q.put(("fatal", "메신저 창을 앞으로 가져오지 못했습니다.\n메신저를 한 번 눌러 앞에 둔 뒤 다시 실행하세요."))
                    return
                notes = []
                if mode == "add":
                    self.event_q.put(("group_status", "상태: 작업용 그룹 준비 중"))
                    st, msg = v3_create_work_group(sess, WORK_GROUP_NAME)
                    print(f"작업용 그룹 만들기: {st} {msg}")
                    if st == "DUPLICATE":
                        if opts.get("exist") == "resume":
                            notes.append(f"기존 '{WORK_GROUP_NAME}' 그룹에 이어서 추가")
                        else:
                            self.event_q.put(("fatal", f"'{WORK_GROUP_NAME}' 그룹이 이미 있어서 멈췄습니다.\n\n"
                                                       "메신저에서 그 그룹 이름을 바꾸거나 지운 뒤 다시 실행하세요.\n"
                                                       "(중간에 멈췄던 작업을 계속하려면 '그 그룹에 이어서 추가'를 고르고 실행)"))
                            return
                    elif st != "OK":
                        self.event_q.put(("fatal", f"작업용 그룹을 만들지 못했습니다: {msg}"))
                        return
                total = len(items)
                completed = True
                for idx, item in enumerate(items, start=1):
                    row = item["row"]
                    self.event_q.put(("group_current", {"row": row, "text": f"현재: {idx}/{total} - {item['label']}"}))
                    self.event_q.put(("group_status", "상태: 검색 중" if mode == "check" else "상태: 검색·추가 중"))
                    print(f"\n######## {idx}/{total} {item['label']} ########")
                    try:
                        person, judg = v3_find_person(sess, item["keywords"])
                        if person is None:
                            status, reason = "실패", judg.get("reason", "검색 실패")
                        elif mode == "check":
                            status, reason = "확인됨", judg.get("person", "")
                        else:
                            sess.guard("내목록 추가")
                            added, why = add_verified_target_to_first_group(person, person["parsed"].get("name", ""))
                            status = "성공" if added else "실패"
                            reason = f"{judg.get('person', '')} · {why}"
                    except ForegroundLost as e:
                        print(f"앞창 잃음: {e}")
                        self.event_q.put(("group_row", {"row": row, "status": "중단", "reason": "메신저가 앞에 없어서 멈춤"}))
                        self.event_q.put(("paused", "메신저 창이 앞에 없어서 멈췄습니다.\n\n실행 중에는 마우스·키보드를 쓰지 마세요.\n"
                                                    "다시 실행하면 '성공' 행은 건너뜁니다.\n(작업용 그룹이 남아 있으니 '그 그룹에 이어서 추가'를 고르세요)"))
                        return
                    except StopRequested:
                        self.event_q.put(("group_row", {"row": row, "status": "중단", "reason": "사용자 중지"}))
                        completed = False
                        break
                    except Exception as e:
                        traceback.print_exc()
                        status, reason = "실패", f"예외 발생: {e.__class__.__name__}"
                    print(f"결과: {status} - {reason}")
                    counts["성공" if status in ("성공", "확인됨") else "실패"] += 1
                    self.event_q.put(("group_row", {"row": row, "status": status, "reason": reason}))

                if mode == "check":
                    done = f"검색 확인 끝: 1명 확인 {counts['성공']}, 실패 {counts['실패']}"
                else:
                    done = f"사용자그룹 추가 끝: 성공 {counts['성공']}, 실패 {counts['실패']}"
                    if not completed:
                        notes.append(f"중지됨. 다음에 '그 그룹에 이어서 추가'를 고르면 남은 사람부터 계속")
                if notes:
                    done += "\n\n- " + "\n- ".join(notes)
                self.event_q.put(("group_done", {"text": done, "rename_notice": mode == "add" and completed and counts["성공"] > 0}))
            except Exception:
                traceback.print_exc()
                self.event_q.put(("fatal", "실행 중 예외가 발생했습니다. logs 폴더의 작업 기록을 확인하세요."))
            finally:
                try:
                    logf.close()
                except Exception:
                    pass

    def _update_group_summary(self):
        s = f = st = 0
        for row in self.group_grid.data:
            n = normalize_result_status(row.get("status", ""))
            if n == "성공":
                s += 1
            elif n == "중단":
                st += 1
            elif n == "실패":
                f += 1
        self.group_summary_var.set(f"성공 {s} / 실패 {f} / 중단 {st}")

    def _poll_events(self):
        try:
            while True:
                event, payload = self.event_q.get_nowait()
                if event == "f12_stop":
                    self.stop()
                    continue
                if event == "group_status":
                    self.group_status_var.set(payload)
                elif event == "group_current":
                    self.group_current_var.set(payload["text"])
                    self.group_grid.set_result(payload["row"], "처리중", "")
                elif event == "group_row":
                    self.group_grid.set_result(payload["row"], payload.get("status", ""), payload.get("reason", ""))
                    self._update_group_summary()
                elif event == "group_done":
                    self.running = False
                    self.group_status_var.set("상태: 대기")
                    self.group_current_var.set("현재: -")
                    self._update_group_summary()
                    if isinstance(payload, dict):
                        messagebox.showinfo("실행 완료", payload.get("text", ""))
                        if payload.get("rename_notice"):
                            messagebox.showwarning(
                                "그룹 이름을 바꿔 주세요",
                                f"'{WORK_GROUP_NAME}' 그룹 이름을 다른 이름으로 바꿔 주세요.\n"
                                "바꾸지 않으면 다음 실행 때 같은 이름이 있어 멈춥니다.\n\n"
                                "[바꾸는 방법]\n"
                                "1. 메신저 [사용자그룹] 탭\n"
                                f"2. '{WORK_GROUP_NAME}' 그룹을 마우스 오른쪽 클릭\n"
                                "3. [그룹 이름 변경] → 새 이름 입력 → Enter")
                    else:
                        messagebox.showinfo("실행 완료", payload)
                elif event == "note_status":
                    self.note_status_var.set(payload)
                elif event == "note_current":
                    self.note_current_var.set(payload["text"])
                    if payload.get("row") is not None:
                        self.note_grid.set_result(payload["row"], "처리중", "")
                elif event == "note_row":
                    self.note_grid.set_result(payload["row"], payload.get("status", ""), payload.get("reason", ""))
                    self._update_note_summary()
                elif event == "note_test_row":
                    self.note_current_var.set("시험: " + payload)
                elif event == "note_done":
                    self.running = False
                    self.note_status_var.set("상태: 대기")
                    self.note_current_var.set("현재: -")
                    self._refresh_note_info()
                    messagebox.showinfo("완료", payload)
                elif event == "paused":
                    self.running = False
                    for v in (self.group_status_var, self.note_status_var):
                        v.set("상태: 멈춤(메신저가 앞에 없음)")
                    self._update_group_summary()
                    self._update_note_summary()
                    messagebox.showwarning("자동으로 멈춤", payload)
                elif event == "fatal":
                    self.running = False
                    self.group_status_var.set("상태: 대기")
                    self.note_status_var.set("상태: 대기")
                    self.group_current_var.set("현재: -")
                    self.note_current_var.set("현재: -")
                    messagebox.showwarning("작업 중단", payload)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_events)

    def _update_note_summary(self):
        if self.note_grid is None:
            return
        success = fail = stop = unclear = 0
        for row in self.note_grid.data:
            norm = normalize_result_status(row.get("status", ""))
            if norm == "성공":
                success += 1
            elif norm == "중단":
                stop += 1
            elif norm == "실패":
                fail += 1
            elif norm == "확인필요":
                unclear += 1
        self.note_summary_var.set(f"성공 {success} / 실패 {fail} / 중단 {stop}" + (f" / 확인필요 {unclear}" if unclear else ""))


def run_ui():
    root = tk.Tk()
    MessengerBatchHelperUIv3(root)
    root.mainloop()


if __name__ == "__main__":
    try:
        run_ui()
    except Exception:
        traceback.print_exc()
