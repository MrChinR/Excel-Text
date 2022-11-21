import dearpygui.dearpygui as dpg
import xlrd
import os


def Read_excel(path, head):
    file = xlrd.open_workbook(path)
    sheet = file.sheets()[0]
    all_rows = sheet.nrows
    all_cols = sheet.ncols
    # print(all_rows, all_cols, row_value)
    data = []
    data1 = []
    for x in range(1, all_rows):
        for y in range(1, all_cols):
            data1 = []
            value = sheet.cell(x, y).value
            if sheet.cell(x, y).value == "":
                continue
            data1.append(value)
            data1.append(head + sheet.cell(x, 0).value + sheet.cell(0, y).value)
            data.append(data1)
    print(data)
    data_sum = []
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
    print(data_sum)
    return data_sum


def Save_in_txt(sender, app_data, user_data):
    try:
        file_path_name = dpg.get_value("file_path_name")
        Head = dpg.get_value("Head")
        net_file_name = dpg.get_value("net_file_name") + "/test.net"
        print(user_data)
        txt_file = open(net_file_name, "w")
        txt_file.write("$CCF" + "\n" + "{" + "\n")
        txt_file.write("\t" + "NET{" + "\n")
        datalist = Read_excel(file_path_name, Head)
        for data in datalist:
            txt_file.write("\t" * 2 + format(data[0], '<15'))
            txt_file.write(":" + ",".join(data[1:]) + ";" + "\n")
        txt_file.write("\t" + "}" + "\n" + "}")
        txt_file.close()
        dpg.configure_item("pop_contents", default_value="変換完了。")
        dpg.configure_item("popup1", label="結果", show=True)
        dpg.configure_item("popup1", show=True)
        if dpg.get_value("open") == True:
            os.startfile(net_file_name, )
    except:
        dpg.configure_item("pop_contents", default_value="エラーがあります。")
        dpg.configure_item("popup1", label="エラー", show=True)


dpg.create_context()
dpg.create_viewport(title='Work_Space', small_icon="ico.ico")
dpg.setup_dearpygui()

with dpg.font_registry():
    with dpg.font(r"C:\Windows\Fonts\BIZ-UDGothicR.ttc", 14,
                  tag="custom font")as default_font:
        dpg.add_font_range_hint(dpg.mvFontRangeHint_Japanese)
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
    dpg.set_value('net_file_name', a)
    print(a)


def _log(sender, app_data, user_data):
    print(f"sender: {sender}, \t app_data: {app_data}, \t user_data: {user_data}")


with dpg.file_dialog(directory_selector=False, show=False, callback=callbackhere, id="file_dialog_id", width=500,
                     height=400):
    dpg.add_file_extension(".xlsx", color=(0, 255, 0, 255), custom_text="[Excel]")

with dpg.file_dialog(directory_selector=True, show=False, callback=callbackhere1, id="folder_dialog_id", width=500,
                     height=400):
    dpg.add_file_extension(".*")

with dpg.window(label="Excel_Matrix-NET_CCF 変換システム(TEST)", width=670, height=250, pos=[200, 200]):
    with dpg.menu_bar():
        with dpg.menu(label="Tools"):
            # dpg.add_menu_item(label="Show About", callback=lambda: dpg.show_tool(dpg.mvTool_About))
            dpg.add_menu_item(label="Show Metrics", callback=lambda: dpg.show_tool(dpg.mvTool_Metrics))
            # dpg.add_menu_item(label="Show Documentation", callback=lambda: dpg.show_tool(dpg.mvTool_Doc))
            dpg.add_menu_item(label="Show Debug", callback=lambda: dpg.show_tool(dpg.mvTool_Debug))
            dpg.add_menu_item(label="Show Style Editor", callback=lambda: dpg.show_tool(dpg.mvTool_Style))
            dpg.add_menu_item(label="Show Font Manager", callback=lambda: dpg.show_tool(dpg.mvTool_Font))
            dpg.add_menu_item(label="Show Item Registry", callback=lambda: dpg.show_tool(dpg.mvTool_ItemRegistry))

        with dpg.menu(label="Settings"):
            dpg.add_menu_item(label="Toggle Fullscreen", callback=lambda: dpg.toggle_viewport_fullscreen())
    with dpg.tab_bar():
        with dpg.tab(label="変換"):
            dpg.add_text("Excel_path_name", tag="file_path_name", color=[255, 0, 0])
            with dpg.tooltip("file_path_name"):
                dpg.add_text("The file path you have selected.")
            with dpg.group(horizontal=True):
                dpg.add_input_text(source='file_path_name')
                dpg.add_button(label="File Selector", callback=lambda: dpg.show_item("file_dialog_id"), width=-1)
            with dpg.group(horizontal=True):
                dpg.add_input_text(tag='net_file_name')
                dpg.add_button(label="Folder Selector", callback=lambda: dpg.show_item("folder_dialog_id"), width=-1)
            with dpg.group(horizontal=True):
                dpg.add_text("接頭詞：")
                dpg.add_input_text(tag="Head", width=255)
                dpg.add_checkbox(label="ファイル開く", tag="open", callback=_log)
                net_file_name = "test.net"
                dpg.add_button(label="変   換", tag="change", callback=Save_in_txt, width=100, height=26)
                dpg.add_button(label="取   消", width=100, height=26, callback=lambda: dpg.stop_dearpygui())
        with dpg.tab(label="Help"):
            dpg.add_text("This is the help page.", color=[255, 0, 0])
     
with dpg.window(label="結果", tag="popup1", show=False, width=150, height=100, pos=[300, 230]):
    dpg.add_text("変換完了。", tag="pop_contents")
    dpg.add_button(label="OK", width=100, height=26, callback=lambda: dpg.configure_item("popup1", show=False))

with dpg.theme() as global_theme:
    with dpg.theme_component(dpg.mvAll):
        dpg.add_theme_color(dpg.mvThemeCol_FrameBg, (54, 64, 60), category=dpg.mvThemeCat_Core)
        dpg.add_theme_style(dpg.mvStyleVar_FrameRounding, 6, category=dpg.mvThemeCat_Core)
        dpg.add_theme_color(dpg.mvThemeCol_Button, (51, 88, 68), category=dpg.mvThemeCat_Core)

    # with dpg.theme_component(dpg.mvInputInt):
    #     dpg.add_theme_color(dpg.mvThemeCol_FrameBg, (51, 88, 68), category=dpg.mvThemeCat_Core)
    #     dpg.add_theme_style(dpg.mvStyleVar_FrameRounding, 6, category=dpg.mvThemeCat_Core)

dpg.bind_theme(global_theme)
# dpg.show_imgui_demo()

dpg.show_viewport()
dpg.start_dearpygui()
dpg.destroy_context()
