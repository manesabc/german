# exe -> pyinstaller german.py german.spec --onefile --windowed --noconsole

if __name__ == '__main__':
    german_version = 1.0


    ################################################################################################
    #
    # Config
    #
    ################################################################################################
    import openpyexcel
    bookname = 'config.xlsx'
    book = openpyexcel.load_workbook(bookname)
    sheet = book['patch10.10']
    config_version = sheet.title
    config = []
    temp = []
    for row in sheet.rows:
        for cell in row:
            temp.append(cell.value)
        config.append(list(temp))
        temp.clear()
    # print(config)
    def get_names(tag, type):
        names = []
        for row in config:
            if row[0] == tag and row[1] == type:
                names.append(row[2])
        return names
    def get_pos(type, name):
        x = 0
        y = 0
        for row in config:
            if row[1] == type and row[2] == name:
                x = row[4]
                y = row[5]
        return x, y
    def get_pos_sub(type, name):
        x = 0
        y = 0
        for row in config:
            if row[1] == type and row[2] == name:
                x = row[6]
                y = row[7]
        return x, y

    ################################################################################################
    #
    # AutoGUI
    #
    ################################################################################################
    import pyautogui
    import win32gui

    window_x = 0
    window_y = 0
    window_loaded = False
    def get_window():
        global window_x, window_y, window_loaded
        parent_handle = win32gui.FindWindow(None, "League of Legends")
        # print(parent_handle)
        if parent_handle > 0:
            win_x1, win_y1, win_x2, win_y2 = win32gui.GetWindowRect(parent_handle)
            window_x = win_x1
            window_y = win_y1
            window_loaded = True
            win32gui.SetForegroundWindow(parent_handle)
        else:
            window_loaded = False
            print("画面の読み込みに失敗しました")
    def left_click(x, y):
        # print(x, y)
        # print(window_loaded)
        # print(window_x, window_y)
        if window_loaded:
            pyautogui.click(window_x + x, window_y + y)
            # print(window_x + x, window_y + y)

    ################################################################################################
    #
    # Layout
    #
    ################################################################################################
    import PySimpleGUI
    layout = [
        [
            PySimpleGUI.Text('メインパス'),
            PySimpleGUI.Radio('栄華', group_id='メインパス', key='栄華', default=True, enable_events=True),
            PySimpleGUI.Radio('覇道', group_id='メインパス', key='覇道', enable_events=True),
            PySimpleGUI.Radio('魔道', group_id='メインパス', key='魔道', enable_events=True),
            PySimpleGUI.Radio('不滅', group_id='メインパス', key='不滅', enable_events=True),
            PySimpleGUI.Radio('天啓', group_id='メインパス', key='天啓', enable_events=True),
        ],
        [
            PySimpleGUI.Text('キーストーン'),
            PySimpleGUI.OptionMenu(values=get_names('栄華', 'キーストーン'), size=(40,1), key='キーストーン')
        ],
        [
            PySimpleGUI.Text('ルーン1'),
            PySimpleGUI.OptionMenu(values=get_names('栄華', 'ルーン1'), size=(40,1), key='ルーン1')
        ],
        [
            PySimpleGUI.Text('ルーン2'),
            PySimpleGUI.OptionMenu(values=get_names('栄華', 'ルーン2'), size=(40,1), key='ルーン2')
        ],
        [
            PySimpleGUI.Text('ルーン3'),
            PySimpleGUI.OptionMenu(values=get_names('栄華', 'ルーン3'), size=(40,1), key='ルーン3')
        ],
        [
            PySimpleGUI.Text('サブパス'),
            PySimpleGUI.Radio('栄華', group_id='サブパス', key='サブ栄華', enable_events=True),
            PySimpleGUI.Radio('覇道', group_id='サブパス', key='サブ覇道', default=True, enable_events=True),
            PySimpleGUI.Radio('魔道', group_id='サブパス', key='サブ魔道', enable_events=True),
            PySimpleGUI.Radio('不滅', group_id='サブパス', key='サブ不滅', enable_events=True),
            PySimpleGUI.Radio('天啓', group_id='サブパス', key='サブ天啓', enable_events=True),
        ],
        [
            PySimpleGUI.Text('ルーン4'),
            PySimpleGUI.OptionMenu(values=get_names('覇道', 'ルーン4'), size=(40,1), key='ルーン4')
        ],
        [
            PySimpleGUI.Text('ルーン5'),
            PySimpleGUI.OptionMenu(values=get_names('覇道', 'ルーン5'), size=(40,1), key='ルーン5')
        ],
        [
            PySimpleGUI.Text('ルーンステータス')
        ],
        [
            PySimpleGUI.Text('ステータス1'),
            PySimpleGUI.OptionMenu(values=get_names('ステータス', 'ステータス1'), size=(40,1), key='ステータス1')
        ],
        [
            PySimpleGUI.Text('ステータス2'),
            PySimpleGUI.OptionMenu(values=get_names('ステータス', 'ステータス2'), size=(40,1), key='ステータス2')
        ],
        [
            PySimpleGUI.Text('ステータス3'),
            PySimpleGUI.OptionMenu(values=get_names('ステータス', 'ステータス3'), size=(40,1), key='ステータス3')
        ],
        [
            PySimpleGUI.Button('実行', key='run')
        ]
    ]


    ################################################################################################
    #
    # Window
    #
    ################################################################################################
    window = PySimpleGUI.Window('german', layout)
    while True:
        event, values = window.read()
        print('event: ', event, ', value: ', values)
        #
        # 終了
        #
        if event in (None, 'quit'):
            break
        #
        # パス設定
        #
        if event == '栄華':
            window['キーストーン'].update(values=get_names('栄華', 'キーストーン'))
            window['ルーン1'].update(values=get_names('栄華', 'ルーン1'))
            window['ルーン2'].update(values=get_names('栄華', 'ルーン2'))
            window['ルーン3'].update(values=get_names('栄華', 'ルーン3'))
        elif event == '覇道':
            window['キーストーン'].update(values=get_names('覇道', 'キーストーン'))
            window['ルーン1'].update(values=get_names('覇道', 'ルーン1'))
            window['ルーン2'].update(values=get_names('覇道', 'ルーン2'))
            window['ルーン3'].update(values=get_names('覇道', 'ルーン3'))
        elif event == '魔道':
            window['キーストーン'].update(values=get_names('魔道', 'キーストーン'))
            window['ルーン1'].update(values=get_names('魔道', 'ルーン1'))
            window['ルーン2'].update(values=get_names('魔道', 'ルーン2'))
            window['ルーン3'].update(values=get_names('魔道', 'ルーン3'))
        elif event == '不滅':
            window['キーストーン'].update(values=get_names('不滅', 'キーストーン'))
            window['ルーン1'].update(values=get_names('不滅', 'ルーン1'))
            window['ルーン2'].update(values=get_names('不滅', 'ルーン2'))
            window['ルーン3'].update(values=get_names('不滅', 'ルーン3'))
        elif event == '天啓':
            window['キーストーン'].update(values=get_names('天啓', 'キーストーン'))
            window['ルーン1'].update(values=get_names('天啓', 'ルーン1'))
            window['ルーン2'].update(values=get_names('天啓', 'ルーン2'))
            window['ルーン3'].update(values=get_names('天啓', 'ルーン3'))
        if event == 'サブ栄華':
            window['ルーン4'].update(values=get_names('栄華', 'ルーン4'))
            window['ルーン5'].update(values=get_names('栄華', 'ルーン5'))
        elif event == 'サブ覇道':
            window['ルーン4'].update(values=get_names('覇道', 'ルーン4'))
            window['ルーン5'].update(values=get_names('覇道', 'ルーン5'))
        elif event == 'サブ魔道':
            window['ルーン4'].update(values=get_names('魔道', 'ルーン4'))
            window['ルーン5'].update(values=get_names('魔道', 'ルーン5'))
        elif event == 'サブ不滅':
            window['ルーン4'].update(values=get_names('不滅', 'ルーン4'))
            window['ルーン5'].update(values=get_names('不滅', 'ルーン5'))
        elif event == 'サブ天啓':
            window['ルーン4'].update(values=get_names('天啓', 'ルーン4'))
            window['ルーン5'].update(values=get_names('天啓', 'ルーン5'))

        #
        # 実行
        #
        if event == 'run':
            get_window()

            count = 0   # サブパスのposにはメインパスに応じて値(2種類のいずれか)をいれる
            mainid = 0
            for mainpath in ['栄華', '覇道', '魔道', '不滅', '天啓']:
                if values[mainpath]:
                    mainid = count
                    x, y = get_pos('メインパス', mainpath)
                    left_click(x, y)
                count += 1

            count = 0
            subid = 0
            for subpath in ['栄華', '覇道', '魔道', '不滅', '天啓']:
                if values['サブ' + subpath]:
                    subid = count
                    if mainid < subid:
                        x, y = get_pos_sub('サブパス', subpath)
                    else:
                        x, y = get_pos('サブパス', subpath)
                    left_click(x, y)
                count += 1

            for item in ['キーストーン', 'ルーン1', 'ルーン2', 'ルーン3', 'ルーン4', 'ルーン5', 'ステータス1', 'ステータス2', 'ステータス3']:
                x, y = get_pos(item, values[item])
                left_click(x, y)

    window.close()

