import os
import win32com.client

# =========================================================
# UNIFY HR - Employee Form v0.1
# Purpose : Create a VBA UserForm layout automatically
# Output  : EmployeeForm_v0_1.xlsm
# =========================================================

OUTPUT_FILE = os.path.abspath("EmployeeForm_v0_1.xlsm")

# MSForms / VBA constants
VBEXT_CT_MSFORM = 3
XL_FILE_FORMAT_XLSM = 52

# Theme
COLOR_FORM_BG = 0xF7F8FA
COLOR_SIDEBAR = 0xF1F5F7
COLOR_HEADER = 0xFFFFFF
COLOR_ACCENT = 0xA8A71F  # BGR-ish COM color approximation for teal; can tune later
COLOR_TEXT = 0x333333
COLOR_MUTED = 0x777777
COLOR_WHITE = 0xFFFFFF
COLOR_BORDER = 0xD9DEE3

FONT_NAME = "Segoe UI"


def set_font(ctrl, size=9, bold=False, color=None):
    try:
        ctrl.Font.Name = FONT_NAME
        ctrl.Font.Size = size
        ctrl.Font.Bold = bold
    except Exception:
        pass
    if color is not None:
        try:
            ctrl.ForeColor = color
        except Exception:
            pass


def add_label(form, name, caption, left, top, width=100, height=18, size=9, bold=False, color=COLOR_TEXT):
    ctrl = form.Controls.Add("Forms.Label.1", name, True)
    ctrl.Caption = caption
    ctrl.Left = left
    ctrl.Top = top
    ctrl.Width = width
    ctrl.Height = height
    ctrl.BackStyle = 0
    set_font(ctrl, size=size, bold=bold, color=color)
    return ctrl


def add_textbox(form, name, left, top, width=170, height=22, text=""):
    ctrl = form.Controls.Add("Forms.TextBox.1", name, True)
    ctrl.Left = left
    ctrl.Top = top
    ctrl.Width = width
    ctrl.Height = height
    ctrl.Text = text
    ctrl.BorderStyle = 1
    set_font(ctrl, size=9, bold=False, color=COLOR_TEXT)
    return ctrl


def add_combobox(form, name, left, top, width=170, height=22):
    ctrl = form.Controls.Add("Forms.ComboBox.1", name, True)
    ctrl.Left = left
    ctrl.Top = top
    ctrl.Width = width
    ctrl.Height = height
    ctrl.Style = 2  # fmStyleDropDownList
    ctrl.BorderStyle = 1
    set_font(ctrl, size=9, bold=False, color=COLOR_TEXT)
    return ctrl


def add_button(form, name, caption, left, top, width=78, height=28, accent=False):
    ctrl = form.Controls.Add("Forms.CommandButton.1", name, True)
    ctrl.Caption = caption
    ctrl.Left = left
    ctrl.Top = top
    ctrl.Width = width
    ctrl.Height = height
    set_font(ctrl, size=9, bold=True if accent else False, color=COLOR_TEXT)
    try:
        if accent:
            ctrl.BackColor = COLOR_ACCENT
            ctrl.ForeColor = COLOR_WHITE
        else:
            ctrl.BackColor = COLOR_WHITE
            ctrl.ForeColor = COLOR_TEXT
    except Exception:
        pass
    return ctrl


def add_frame(form, name, caption, left, top, width, height, bg=None):
    ctrl = form.Controls.Add("Forms.Frame.1", name, True)
    ctrl.Caption = caption
    ctrl.Left = left
    ctrl.Top = top
    ctrl.Width = width
    ctrl.Height = height
    set_font(ctrl, size=9, bold=False, color=COLOR_MUTED)
    if bg is not None:
        try:
            ctrl.BackColor = bg
        except Exception:
            pass
    return ctrl


def add_panel_label(form, name, left, top, width, height, back_color):
    # MSForms has no true panel; Label with SpecialEffect works as background block
    ctrl = form.Controls.Add("Forms.Label.1", name, True)
    ctrl.Caption = ""
    ctrl.Left = left
    ctrl.Top = top
    ctrl.Width = width
    ctrl.Height = height
    ctrl.BackStyle = 1
    ctrl.BackColor = back_color
    ctrl.SpecialEffect = 0
    return ctrl


def build_employee_form():
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True

    wb = excel.Workbooks.Add()
    vbproj = wb.VBProject

    component = vbproj.VBComponents.Add(VBEXT_CT_MSFORM)
    component.Name = "frmEmployee"
    form = component.Designer

    # NOTE: Some environments do not allow direct form Width/Height setting through Designer.
    # Controls are laid out within default form size; manual resize can be done later if needed.
    try:
        form.Caption = "UNIFY HR - Employee Registration"
        form.BackColor = COLOR_FORM_BG
    except Exception:
        pass

    # Layout constants
    sidebar_x = 12
    sidebar_y = 12
    sidebar_w = 112
    sidebar_h = 300

    content_x = 138
    content_y = 18

    label_w = 70
    input_w = 162
    row_h = 34
    field_h = 22

    right_x = 410
    right_y = 54

    # Background blocks
    add_panel_label(form, "bgSidebar", sidebar_x, sidebar_y, sidebar_w, sidebar_h, COLOR_SIDEBAR)
    add_panel_label(form, "bgHeader", 0, 0, 620, 44, COLOR_HEADER)

    # Header
    add_label(form, "lblBrand", "UNIFY HR", 18, 12, 90, 20, size=11, bold=True, color=COLOR_ACCENT)
    add_label(form, "lblTitle", "Employee Registration", content_x, 13, 210, 22, size=14, bold=True, color=COLOR_TEXT)
    add_label(form, "lblSubTitle", "Create and manage employee master data", content_x, 34, 240, 14, size=8, color=COLOR_MUTED)

    # Sidebar menu
    menu_items = [
        ("기본 정보", True),
        ("조직 정보", False),
        ("근무 정보", False),
        ("연락처", False),
        ("급여 정보", False),
        ("메모", False),
    ]
    y = 60
    for idx, (caption, active) in enumerate(menu_items):
        if active:
            add_panel_label(form, f"navActive{idx}", sidebar_x + 8, y - 4, 92, 26, COLOR_WHITE)
            add_panel_label(form, f"navBar{idx}", sidebar_x + 8, y - 4, 4, 26, COLOR_ACCENT)
            add_label(form, f"nav{idx}", caption, sidebar_x + 18, y + 1, 80, 16, size=9, bold=True, color=COLOR_TEXT)
        else:
            add_label(form, f"nav{idx}", caption, sidebar_x + 18, y + 1, 80, 16, size=9, bold=False, color=COLOR_MUTED)
        y += 34

    # Section title
    add_label(form, "lblSectionBasic", "기본 정보", content_x, 62, 120, 20, size=11, bold=True, color=COLOR_TEXT)
    add_label(form, "lblSectionHint", "사원 마스터 등록에 필요한 기본 항목입니다.", content_x, 82, 260, 16, size=8, color=COLOR_MUTED)

    # Fields left column
    start_y = 112
    fields_left = [
        ("사번", "txtEmployeeID", "EMP000001", "text"),
        ("성명", "txtName", "", "text"),
        ("영문명", "txtEnglishName", "", "text"),
        ("생년월일", "txtBirthDate", "YYYY-MM-DD", "text"),
        ("입사일", "txtHireDate", "YYYY-MM-DD", "text"),
    ]

    y = start_y
    for label, name, default, kind in fields_left:
        add_label(form, f"lbl_{name}", label, content_x, y + 4, label_w, 16, size=9, bold=False, color=COLOR_MUTED)
        tb = add_textbox(form, name, content_x + 78, y, input_w, field_h, default)
        if name == "txtEmployeeID":
            try:
                tb.Enabled = False
                tb.BackColor = 0xEEEEEE
            except Exception:
                pass
        y += row_h

    # Fields right-middle column
    mid_x = content_x + 270
    y = start_y
    fields_right = [
        ("부서", "cboDepartment"),
        ("직책", "cboPosition"),
        ("직급", "cboGrade"),
        ("고용형태", "cboEmploymentType"),
        ("재직구분", "cboWorkStatus"),
    ]

    for label, name in fields_right:
        add_label(form, f"lbl_{name}", label, mid_x, y + 4, label_w, 16, size=9, color=COLOR_MUTED)
        cb = add_combobox(form, name, mid_x + 78, y, input_w, field_h)
        # sample list for visual check only
        try:
            if name == "cboDepartment":
                cb.AddItem("경영지원팀")
                cb.AddItem("생산팀")
                cb.AddItem("품질팀")
            elif name == "cboPosition":
                cb.AddItem("팀장")
                cb.AddItem("파트장")
                cb.AddItem("담당")
            elif name == "cboGrade":
                cb.AddItem("사원")
                cb.AddItem("대리")
                cb.AddItem("과장")
            elif name == "cboEmploymentType":
                cb.AddItem("정규직")
                cb.AddItem("계약직")
                cb.AddItem("파견직")
            elif name == "cboWorkStatus":
                cb.AddItem("재직")
                cb.AddItem("휴직")
                cb.AddItem("퇴직")
        except Exception:
            pass
        y += row_h

    # Photo panel
    add_frame(form, "fraPhoto", "사진", right_x, 230, 120, 92, bg=COLOR_WHITE)
    add_label(form, "lblPhotoIcon", "PHOTO", right_x + 37, 267, 56, 18, size=10, bold=True, color=COLOR_MUTED)
    add_button(form, "cmdPhoto", "사진 등록", right_x + 18, 327, 82, 24, accent=False)

    # Memo area
    add_label(form, "lblMemo", "메모", content_x, 292, 60, 16, size=9, color=COLOR_MUTED)
    memo = form.Controls.Add("Forms.TextBox.1", "txtMemo", True)
    memo.Left = content_x + 78
    memo.Top = 286
    memo.Width = 250
    memo.Height = 58
    memo.MultiLine = True
    memo.BorderStyle = 1
    set_font(memo, size=9, color=COLOR_TEXT)

    # Bottom line and buttons
    add_panel_label(form, "bottomLine", 0, 365, 620, 1, COLOR_BORDER)
    add_button(form, "cmdReset", "초기화", 258, 382, 72, 28, accent=False)
    add_button(form, "cmdSave", "저장", 338, 382, 72, 28, accent=True)
    add_button(form, "cmdSaveNew", "저장 후 신규", 418, 382, 90, 28, accent=False)
    add_button(form, "cmdClose", "닫기", 516, 382, 72, 28, accent=False)

    # Create a starter module to open the form
    module = vbproj.VBComponents.Add(1)  # standard module
    module.Name = "modOpenEmployeeForm"
    module.CodeModule.AddFromString(
        'Option Explicit\n\n'
        'Public Sub OpenEmployeeForm()\n'
        '    frmEmployee.Show\n'
        'End Sub\n'
    )

    # Save
    save_path = os.path.abspath("EmployeeForm_v0_1.xlsm")
    wb.SaveAs(save_path, FileFormat=XL_FILE_FORMAT_XLSM)

    print("Employee Form v0.1 생성 완료")
    print("저장 위치:", save_path)
    print("Excel에서 Alt+F11 > frmEmployee 를 열어 확인하세요.")
    input("확인 후 Enter를 누르면 Excel을 종료합니다.")

    wb.Close(False)
    excel.Quit()


if __name__ == "__main__":
    build_employee_form()
