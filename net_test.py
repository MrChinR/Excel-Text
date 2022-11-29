import dearpygui.dearpygui as dpg
import xlrd
import os


def f_to_i(value):
    if isinstance(value, float) and value % 1 == 0.0:
        value = str(int(value))
    return value


def Read_excel(path, head, head1):
    file = xlrd.open_workbook(path)
    sheet = file.sheets()[0]
    all_rows = sheet.nrows
    all_cols = sheet.ncols
    print(all_rows, all_cols)
    Column_BF = dpg.get_value("Column B/F") - 1
    Column_BTM = dpg.get_value("Column BTM") - 1
    Column_NET = dpg.get_value("Column NET") - 1
    Row_Start = dpg.get_value("Row Start") - 1
    NC_name = dpg.get_value("NC")
    i = 1
    data_sum = []
    if dpg.get_value("mode") == "単独エクセルモード":
        data = []
        for x in range(Row_Start, all_rows):
            data1 = []
            value = f_to_i(sheet.cell(x, Column_NET).value)
            value1 = f_to_i(sheet.cell(x, Column_BF).value)
            value2 = f_to_i(sheet.cell(x, Column_BTM).value)
            if value == "":
                continue
            if value == NC_name:
                value = value + f"<{i}>"
                i = i + 1
            data1.append(value)
            if value1 != "":
                data1.append(head1 + str(value1))
            if value2 != "":
                data1.append(head + str(value2))
            data.append(data1)

    elif dpg.get_value("mode") == "背面Matrixモード":
        # print(all_rows, all_cols, row_value)
        data = []
        for x in range(1, all_rows):
            for y in range(1, all_cols):
                data1 = []
                value = f_to_i(sheet.cell(x, y).value)
                if sheet.cell(x, y).value == "":
                    continue
                if value == NC_name:
                    value = value + f"<{i}>"
                    i = i + 1
                data1.append(value)
                data1.append(head + f_to_i(sheet.cell(x, 0).value) + f_to_i(sheet.cell(0, y).value))
                data.append(data1)
        # print(data)
    else:
        sheet1 = file.sheets()[1]
        all_rows1 = sheet1.nrows
        all_cols1 = sheet1.ncols
        print(all_rows1, all_cols1)
        data = []
        for x in range(Row_Start, all_rows):
            data1 = []
            value = f_to_i(sheet.cell(x, Column_NET).value)
            value1 = f_to_i(sheet.cell(x, Column_BF).value)
            value2 = f_to_i(sheet.cell(x, Column_BTM).value)
            if value == "":
                continue
            if value == NC_name:
                value = value + f"<{i}>"
                i = i + 1
            data1.append(value)
            if value1 != "":
                data1.append(head1 + str(value1))
            if value2 != "":
                data1.append(head + str(value2))
            data.append(data1)
        for x in range(1, all_rows1):
            for y in range(1, all_cols1):
                data1 = []
                value = f_to_i(sheet1.cell(x, y).value)
                if sheet1.cell(x, y).value == "":
                    continue
                if value == NC_name:
                    value = value + f"<{i}>"
                    i = i + 1
                data1.append(value)
                data1.append(head + f_to_i(sheet1.cell(x, 0).value) + f_to_i(sheet1.cell(0, y).value))
                data.append(data1)
        # print(data)
    for m in data:
        flag = 0
        a = 0
        for n in data_sum:
            if n[0] == m[0]:
                data_sum[a].append(m[1])
                flag = 1
                break
            a = a + 1
        if flag == 1:
            continue
        else:
            data_sum.append(m)
    # print(data_sum)
    return data_sum


def Save_in_txt(sender, app_data, user_data):
    try:
        mouse_pos = dpg.get_mouse_pos(local=False)
        dpg.configure_item("popup1", pos=mouse_pos)
        print(mouse_pos)
        file_path_name = dpg.get_value("file_path_name")
        Head = dpg.get_value("Head")
        Head1 = dpg.get_value("Head1")
        net_name = dpg.get_value("net_file_name")
        if net_name == "Net_file_name":
            net_file_name = dpg.get_value("net_path_name") + "/net.net"
        else:
            net_file_name = dpg.get_value("net_path_name") + "/" + net_name
        txt_file = open(net_file_name, "w")
        txt_file.write("$CCF" + "\n" + "{" + "\n")
        txt_file.write("\t" + "/*" + "\n")
        txt_file.write("\t" + " " * 3 + "Excel File path : " + file_path_name + "\n")
        txt_file.write("\t" + "/*" + "\n")
        txt_file.write("\t" + "NET{" + "\n")
        datalist = Read_excel(file_path_name, Head, Head1)
        maxlength = max(len(s[0]) for s in datalist)
        for data in datalist:
            txt_file.write("\t" * 2 + format(data[0], f'<{maxlength + 1}') + ": ")
            f = 0
            for i in data[1:-1]:
                f = f + len(i)
                if f > 50:
                    txt_file.write("\n" + "\t" * 2 + " " * (maxlength + 3))
                    f = 0
                txt_file.write(i + ",")
            txt_file.write(data[-1] + ";" + "\n")
            # txt_file.write(":" + ",".join(data[1:]) + ";" + "\n")
        txt_file.write("\t" + "}" + "\n" + "}")
        txt_file.close()
        dpg.configure_item("pop_contents", default_value="変換完了。")
        dpg.configure_item("popup1", label="結果", show=True)
        dpg.configure_item("popup1", show=True)
        if dpg.get_value("open") == True:
            os.startfile(net_file_name, )
    except Exception as e:
        print(e)
        dpg.configure_item("pop_contents", default_value="エラーがあります。")
        dpg.configure_item("popup1", label="エラー", show=True)


dpg.create_context()
dpg.create_viewport(title='Work_Space', small_icon="ico.ico")
dpg.setup_dearpygui()

with dpg.font_registry():
    # with dpg.font(r"C:\Windows\Fonts\BIZ-UDGothicB.ttc", 14,
    #               tag="custom font"):
    #     dpg.add_font_range_hint(dpg.mvFontRangeHint_Japanese)
    with dpg.font(r"C:\Windows\Fonts\UDDigiKyokashoN-B.ttc", 16,
                  tag="custom font B")as default_font:
        dpg.add_font_range_hint(dpg.mvFontRangeHint_Japanese)
    # with dpg.font(r"C:\Windows\Fonts\msyhbd.ttc", 17,
    #               tag="custom font Chinese")as default_font:
    #     dpg.add_font_range_hint(dpg.mvFontRangeHint_Chinese_Full)
    # dpg.bind_font(dpg.last_container())
dpg.bind_font(default_font)  # Binds the font globally



def callbackhere(sender, app_data, user_data):
    print("Sender: ", sender)
    print("App Data: ", app_data)
    a = app_data['file_path_name']
    dpg.set_value('file_path_name', a)
    print(a)


def callbackhere1(sender, app_data, user_data):
    print("Sender: ", sender)
    print("App Data: ", app_data)
    a = app_data['file_path_name']
    dpg.set_value('net_path_name', a)
    print(a)


def _log(sender, app_data, user_data):
    print(f"sender: {sender}, \t app_data: {app_data}, \t user_data: {user_data}")


def _help(message):
    last_item = dpg.last_item()
    group = dpg.add_group(horizontal=True)
    dpg.move_item(last_item, parent=group)
    dpg.capture_next_item(lambda s: dpg.move_item(s, parent=group))
    t = dpg.add_text("(?)", color=[0, 255, 0])
    with dpg.tooltip(t):
        dpg.add_text(message)


with dpg.file_dialog(directory_selector=False, show=False, callback=callbackhere, id="file_dialog_id", width=500,
                     height=400):
    dpg.add_file_extension(".xlsx,.xls", color=(0, 255, 0, 255), custom_text="Excel")

with dpg.file_dialog(directory_selector=True, show=False, callback=callbackhere1, id="folder_dialog_id", width=500,
                     height=400):
    dpg.add_file_extension(".*")

with dpg.window(label="Excel-NET_CCF 変換システム(TEST)", width=670, height=230, pos=[200, 200], tag="main_window",
                no_resize=True):
    with dpg.menu_bar():
        with dpg.menu(label="ツール"):
            # dpg.add_menu_item(label="Show About", callback=lambda: dpg.show_tool(dpg.mvTool_About))
            dpg.add_menu_item(label="Show Metrics", callback=lambda: dpg.show_tool(dpg.mvTool_Metrics))
            # dpg.add_menu_item(label="Show Documentation", callback=lambda: dpg.show_tool(dpg.mvTool_Doc))
            dpg.add_menu_item(label="Show Debug", callback=lambda: dpg.show_tool(dpg.mvTool_Debug))
            dpg.add_menu_item(label="Show Style Editor", callback=lambda: dpg.show_tool(dpg.mvTool_Style))
            dpg.add_menu_item(label="Show Font Manager", callback=lambda: dpg.show_tool(dpg.mvTool_Font))
            dpg.add_menu_item(label="Show Item Registry", callback=lambda: dpg.show_tool(dpg.mvTool_ItemRegistry))

        with dpg.menu(label="設定"):
            dpg.add_menu_item(label="フルスクリーン", callback=lambda: dpg.toggle_viewport_fullscreen())
    with dpg.tab_bar():
        with dpg.tab(label="変換"):
            with dpg.group(horizontal=True):
                dpg.add_radio_button(("単独エクセルモード", "背面Matrixモード", "BF+背面Matrixモード"), default_value="単独エクセルモード",
                                     tag="mode",
                                     callback=_log, horizontal=True)
                dpg.add_text("NC端子名：")
                dpg.add_input_text(tag="NC", width=-1)
            with dpg.group(horizontal=True):
                dpg.add_input_text(default_value="Excel_path_name", tag='file_path_name')
                dpg.add_button(label="ファイル参照", callback=lambda: dpg.show_item("file_dialog_id"), width=-1)
            with dpg.group(horizontal=True):
                dpg.add_input_text(default_value="Net_path_name", tag='net_path_name', width=300)
                dpg.add_input_text(default_value="Net_file_name", tag='net_file_name', width=125)
                dpg.add_button(label="フォルダー参照", callback=lambda: dpg.show_item("folder_dialog_id"), width=-1)
            with dpg.group(horizontal=True):
                dpg.add_text("BF接頭詞：")
                dpg.add_input_text(tag="Head1", width=66)
                dpg.add_text("背面接頭詞：")
                dpg.add_input_text(tag="Head", width=66)
                dpg.add_checkbox(label="ファイル開く", tag="open", callback=_log)
                net_file_name = "test.net"
                dpg.add_button(label="変   換", tag="change", callback=Save_in_txt, width=101, height=26)
                dpg.add_button(label="取   消", tag="cancel", width=101, height=26, callback=lambda: dpg.stop_dearpygui())
            with dpg.tooltip("cancel"):
                dpg.add_text("よろしいですか？")
        with dpg.tab(label="Excel行列指定"):
            dpg.add_slider_int(label="B/F列", tag="Column B/F", max_value=10, callback=_log)
            _help("CTRL+clickで手入力.以降同様")
            dpg.add_slider_int(label="背面ﾊﾟｯﾄﾞ列", tag="Column BTM", max_value=10, callback=_log)
            dpg.add_slider_int(label="ネット名列", tag="Column NET", max_value=10, callback=_log)
            dpg.add_slider_int(label="内容開始行", tag="Row Start", max_value=10, callback=_log)
        with dpg.tab(label="説明"):
            dpg.add_text("単独エクセルモードとは、一つのシートに全てのネット情報が含まれている事。事前にExcelを調整する必要がある。",
                         color=[255, 0, 0], wrap=650)
            dpg.add_text("BF+背面Matrixモードとはシート1にBF情報、シート2に背面Matrixネット情報各々含まれている事。",
                         color=[255, 0, 0], wrap=650)
            dpg.add_text("除外端子名について、客先NC端子名が同じの場合、端子名を入力してください。",
                         color=[255, 0, 0], wrap=650)
            dpg.add_text("netファイルの名前を指定しない場合は、net.netとする。",
                         color=[255, 0, 0], wrap=650)
            dpg.add_text("Excelの列を指定する必要がある。Excel行列指定のtabで指定を行う。",
                         color=[255, 0, 0], wrap=650)

with dpg.window(label="結果", tag="popup1", show=False, modal=True, autosize=True):
    dpg.add_text("変換完了。", tag="pop_contents")
    dpg.add_button(label="OK", width=100, height=26, callback=lambda: dpg.configure_item("popup1", show=False))

with dpg.theme() as global_theme:
    with dpg.theme_component(dpg.mvAll):
        dpg.add_theme_color(dpg.mvThemeCol_FrameBg, (54, 64, 60), category=dpg.mvThemeCat_Core)
        dpg.add_theme_style(dpg.mvStyleVar_FrameRounding, 6, category=dpg.mvThemeCat_Core)
        dpg.add_theme_style(dpg.mvStyleVar_WindowRounding, 5, category=dpg.mvThemeCat_Core)
        dpg.add_theme_style(dpg.mvStyleVar_GrabRounding, 6, category=dpg.mvThemeCat_Core)
        dpg.add_theme_style(dpg.mvStyleVar_WindowTitleAlign, 0.5, 0.5, category=dpg.mvThemeCat_Core)
        dpg.add_theme_style(dpg.mvStyleVar_FrameBorderSize, 0.1, category=dpg.mvThemeCat_Core)
        dpg.add_theme_style(dpg.mvStyleVar_ItemSpacing, 9, 8, category=dpg.mvThemeCat_Core)
        dpg.add_theme_color(dpg.mvThemeCol_Button, (51, 88, 68), category=dpg.mvThemeCat_Core)

    # with dpg.theme_component(dpg.mvInputInt):
    #     dpg.add_theme_color(dpg.mvThemeCol_FrameBg, (51, 88, 68), category=dpg.mvThemeCat_Core)
    #     dpg.add_theme_style(dpg.mvStyleVar_FrameRounding, 6, category=dpg.mvThemeCat_Core)
dpg.bind_theme(global_theme)

# dpg.set_primary_window("main_window", True)
dpg.show_viewport()
dpg.start_dearpygui()
dpg.destroy_context()
