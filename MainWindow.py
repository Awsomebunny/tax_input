import PySimpleGUI as sg
import datetime
import time
import FormXLS
import locale
import json

locale.setlocale(locale.LC_ALL, ('ru_RU', 'UTF-8'))
sg.theme_global("SystemDefaultForReal")
digits = ("0", "1", "2", "3", "4", "5", "6", "7", "8", "9")
def datennumber_frame_layout():
    for i in range(20):
        yield [[sg.Text('Дата'), sg.Frame(title="", layout=[
            [sg.InputText("ДД.ММ.ГГГГ", enable_events=True, key="-Date[0," + str(i) + "]-", disabled=True),
             sg.CalendarButton("Выбрать", format="%d.%m.%Y", begin_at_sunday_plus=1)]],
                                          key="InputFrame[0," + str(i) + "]"),
                sg.Text('Номер'), sg.InputText(enable_events=True, key="-Number[0," + str(i) + "]-")]]


def daten_number_YN_frame_layout():
    for i in range(20):
        yield [[sg.Text('Дата письма'), sg.InputText("ДД.ММ.ГГГГ", enable_events=True, key="-Date[1," + str(i) + "]-"),
                sg.CalendarButton("Выбрать", format="%d.%m.%Y", begin_at_sunday_plus=1),
                sg.Text('Номер письма'), sg.InputText(enable_events=True, key="-Number[1," + str(i) + "]-"),
                sg.Text("Поддержано управлением"),
                sg.Combo(["Да", "Нет"], enable_events=True, key="-Combo[1," + str(i) + "]-")]]


def dates_frame_layout():
    for i in range(20):
        yield [[sg.Text('Дата'), sg.InputText("ДД.ММ.ГГГГ", key="-Date[2," + str(i) + "]-", enable_events=True),
                sg.CalendarButton("Выбрать", format="%d.%m.%Y", begin_at_sunday_plus=1),
                sg.Text('Номер'), sg.InputText(enable_events=True, key="-Number[2," + str(i) + "]-")],
               [sg.Text('Дата вручения'),
                sg.InputText("ДД.ММ.ГГГГ", key="-ArrivalDate[2," + str(i) + "]-", enable_events=True),
                sg.CalendarButton("Выбрать", format="%d.%m.%Y", begin_at_sunday_plus=1),
                sg.Text("Способ вручения"),
                sg.Combo(["Лично", "ТКС", "Почта"], key="-Combo[2," + str(i) + "]-", enable_events=True)]]


def contragent_frame_layout():
    return [[sg.Text('Наименование'), sg.InputText(key="-ContragentName-"), sg.Text('ИНН'),
             sg.InputText(key="ContragentNumber")]]


declaration_list = []
contragent_list = []
corrected_values = []
datennumber = datennumber_frame_layout()
datennumberYN = daten_number_YN_frame_layout()
dates = dates_frame_layout()
declaration_frame_layout = [
    [sg.Text('Регистрационный номер декларации'), sg.InputText(enable_events=True, key="-DeclarationNumber-")],
    [sg.Text('Период, за который представлена НД'),
     sg.Combo([str(i) for i in range(1, 5)], enable_events=True, key="-DeclarationPeriodCombo-"),
     sg.Text('квартал'),
     sg.Spin([i for i in range(2000, int(datetime.datetime.now().year) + 1)],
             initial_value=str(datetime.datetime.now().year), enable_events=True,
             key="-DeclarationYearSpin-"), sg.Text('года')],
    [sg.Text("Декларация актуальна"),
     sg.Combo(["Да", "Нет"], enable_events=True, key="-ActualDeclarationCombo-")],
    [sg.Button("Добавить декларацию")]
]
menu_def = [['Файл', ['Открыть', 'Сохранить', 'Сохранить как...', ]],
            ['Помощь', 'О программе'],
            ['Выйти']]
layout = [
    [sg.Menu(menu_def)],
    [sg.Text('Код ИФНС'), sg.InputText(enable_events=True, key="-OfficeNumber-")
     ],
    [sg.Text('Наименование ВП'), sg.InputText(enable_events=True, key="-CriminalName-"), sg.Text("Тип ВП"),
     sg.Combo(["ИП (12 цифр ИНН)", "Юр. лицо (10 цифр ИНН)"], key="-CriminalTypeCombo-", enable_events=True,
              background_color="orange")
     ],
    [sg.Text('ИНН ВП'), sg.InputText(enable_events=True, key="-CriminalNumber-", disabled=True), sg.Button("Поиск" , disabled=True)
     ],
    [sg.Text('ВП акцептован'), sg.Combo(["Да", "Нет"], enable_events=True, key="-ChallengeAcceptedCombo-")
     ],
    [sg.Text('Номер письма об акцептовании'), sg.InputText(enable_events=True, key="-ChallengeNumber-"),
     sg.Text('Дата письма об акцептовании'),
     sg.InputText("ДД.ММ.ГГГГ", enable_events=True, key="-ChallengeDate-"),
     sg.CalendarButton("Выбрать", format="%d.%m.%Y", begin_at_sunday_plus=1)
     ],
    [sg.Frame('Сведения о налоговой декларации', declaration_frame_layout, font='Any 12', title_color='blue',
              key="-DeclarationFrame-")] + [sg.Frame("Сведения о поданных налоговых декларациях", [[sg.Table(
        values=[["-------------", "-------------", "-------------"]],
        headings=["Рег.номер", "Налоговый период", "Актуальная"], max_col_width=25,

        background_color='lightyellow',
        auto_size_columns=True,
        vertical_scroll_only=True,
        display_row_numbers=True,
        justification='right',
        num_rows=5,
        alternating_row_color='lightyellow',
        key='-DeclarationsTable-',
        row_height=35,
        tooltip='This is a table')],
        [sg.Text(
            "Кол-во элементов:     " + str(
                len(
                    declaration_list)),
            enable_events=True,
            key="-DeclarationsQuantity-",
            auto_size_text=True)]],
                                                     font='Any 12', title_color='blue')],
    [sg.Text('Дата начала КНП'), sg.InputText("ДД.ММ.ГГГГ", enable_events=True, key="-InspectionStartDate-"),
     sg.CalendarButton("Выбрать", format="%d.%m.%Y", begin_at_sunday_plus=1)
     ],
    [sg.Text('Дата окончания 2-х месячного срока КНП'),
     sg.InputText("ДД.ММ.ГГГГ", enable_events=True, key="-InspectionStartDate+2-")
     ],
    [sg.Text('Дата окончания 3-х месячного срока КНП'),
     sg.InputText("ДД.ММ.ГГГГ", enable_events=True, key="-InspectionStartDate+3-")
     ],
    [sg.Text('Сумма сложного расхождения дошедшая до ВП, отрабатываемая в ходе КНП, тыс. руб. (1-НР)'),
     sg.InputText(enable_events=True, key="-FraudSum-")
     ],
    [sg.Text('Планируемая сумма доначислений по результатам КНП всего, тыс.руб.'),
     sg.InputText(enable_events=True, key="-CompensationSum-")
     ],
    [sg.Frame('Сведения о проблемном контрагенте, установленном в рамках "сложного" расхождения',
              [[sg.Text('Наименование'), sg.InputText(key="-MainContragentName-"), sg.Text("Тип"),
                sg.Combo(["ИП (12 цифр ИНН)", "Юр. лицо (10 цифр ИНН)"], key="-MainContragentTypeCombo-",
                         enable_events=True), sg.Text('ИНН'),
                sg.InputText(key="-MainContragentNumber-", disabled=True, enable_events=True)]], font='Any 12',
              title_color='blue')],
    [sg.Frame('Замена проблемного контрагента на иное ЮЛ', [
        [sg.Text('Наименование'), sg.InputText(key="-ContragentName-", enable_events=True), sg.Text("Тип"),
         sg.Combo(["ИП (12 цифр ИНН)", "Юр. лицо (10 цифр ИНН)"], key="-ContragentTypeCombo-", enable_events=True),
         sg.Text('ИНН'),
         sg.InputText(key="-ContragentNumber-", enable_events=True, disabled=True)]] + [[sg.Button("Добавить")]],
              font='Any 12',
              title_color='blue')],
    [sg.Frame("Сведения о заменах проблемного контрагента", [[sg.Table(
        values=[["--------------------------", "--------------------------"]], headings=["ИНН", "Наименование"],
        max_col_width=25,

        background_color='lightyellow',
        auto_size_columns=True,
        vertical_scroll_only=False,
        display_row_numbers=True,
        justification='right',
        num_rows=5,
        alternating_row_color='lightyellow',
        key='-ContragentsTable-',
        row_height=35,
        tooltip='This is a table')],
        [sg.Text(
            "Кол-во элементов: 000",
            enable_events=True, key="-ContragentsQuantity-",
            auto_size_text=True)]], font='Any 12',
              title_color='blue')],
    [sg.Frame('Письмо ТНО в УФНС о согласовании продления КНП до 3 -х месяцев', next(datennumber),
              font='Any 12', title_color='blue')],
    [sg.Frame('Письмо УФНС о согласовании срока продления КНП до 3 -х месяцев', next(datennumber),
              font='Any 12', title_color='blue')],
    [sg.Frame('Решение о продлении КНП до 3-х месяцев', next(datennumber), font='Any 12', title_color='blue')],
    [sg.Frame('Направление проекта Акта КНП в УФНС', next(datennumber), font='Any 12', title_color='blue')],
    [sg.Frame('Результат рассмотрения проекта  Акта КНП Управлением ', next(datennumberYN), font='Any 12',
              title_color='blue')],
    [sg.Frame('Акт КНП', next(dates), font='Any 12', title_color='blue')],
    [sg.Frame('Возражения на АКТ КНП от НП', [
        [sg.Text('Дата представления возражений НП'), sg.InputText("ДД.ММ.ГГГГ", enable_events=True, key="-ClaimDate-"),
         sg.CalendarButton("Выбрать", format="%d.%m.%Y", begin_at_sunday_plus=1),
         sg.Text('Номер'), sg.InputText(enable_events=True, key="-ClaimNumber-")],
        [sg.Text('Дата протокола заседания рабочей группы УФНС по ЧО по вопросу рассмотрения акта КНП и возражений НП'),
         sg.InputText("ДД.ММ.ГГГГ", enable_events=True, key="-ClaimProcessedDate-"),
         sg.CalendarButton("Выбрать", format="%d.%m.%Y")],
        [sg.Text("Итоги рассмотрения"), sg.Combo(
            ["Решение об отказе в привлечении к ответственности", "Решение о проведении дополнительных мероприятий"],
            key="-DecisionCombo-", enable_events=True)]], font='Any 12', title_color='blue')],
    [sg.Frame('Решение о проведении дополнительных мероприятий', next(datennumber), font='Any 12',
              title_color='blue')],
    [sg.Frame('Дополнение к Акту КНП', next(dates), font='Any 12', title_color='blue')],
    [sg.Frame('Возражения на дополнение к АКТУ КНП от НП', [
        [sg.Text('Дата представления возражений НП'),
         sg.InputText("ДД.ММ.ГГГГ", enable_events=True, key="-PretensionDate-"),
         sg.CalendarButton("Выбрать", format="%d.%m.%Y", begin_at_sunday_plus=1),
         sg.Text('Номер'), sg.InputText(key="-PretensionNumber-")],
        [sg.Text(
            'Дата протокола заседания рабочей группы УФНС по ЧО по вопросу рассмотрения дополнения акта КНП и возражений НП'),
            sg.InputText("ДД.ММ.ГГГГ", enable_events=True, key="-AdditionProcessedDate-"),
            sg.CalendarButton("Выбрать", format="%d.%m.%Y", begin_at_sunday_plus=1)],
        [sg.Text("Итоги рассмотрения"),
         sg.Combo(["Решение об отказе в привлечении к ответственности", "Позиция ТНО поддержана"],
                  key="-AdditionalDecisionCombo-",
                  enable_events=True)]], font='Any 12', title_color='blue')],

    [sg.Frame('Направление проекта Решения КНП в УФНС', [[sg.Frame('Сведения о письме', next(datennumber),
                                                                   font='Any 12', title_color='blue')]],
              font='Any 12', title_color='blue')],
    [sg.Frame('Решение КНП', next(dates) + [[sg.Text("Статус решения"), sg.Combo(
        ["Решение об отказе", "Решение о привлечении к ответственности"], key="-FinalDecisionCombo-",
        enable_events=True)]],
              font='Any 12', title_color='blue')],
    [sg.Frame('Направление проекта Решения о принятии обеспечительных мер по результатам КНП в УФНС',
              next(datennumber),
              font='Any 12', title_color='blue')],
    [sg.Frame('Результат рассмотрения проекта Решения о принятии обеспечительных мер по результатам КНП',
              next(datennumberYN), font='Any 12',
              title_color='blue')],
    [sg.Frame('Решение о принятии обеспечительных мер по результатам КНП', next(dates), font='Any 12',
              title_color='blue')],
    [sg.Text("Меры взыскания"), sg.InputText(enable_events=True)],
    [sg.Frame('Направление в УФНС проекта сообщения о преступлении (ст. 32 НК РФ)', next(datennumber),
              font='Any 12', title_color='blue')],
    [sg.Frame('Направление проекта сообщения о преступлении  (ст. 82 НК РФ)', next(datennumber),
              font='Any 12', title_color='blue')],
    [sg.Frame('Cлужебная записка передачи материалов  в ОПАИД', next(datennumber),
              font='Any 12', title_color='blue')],
    [sg.Text("УД возбуждено"), sg.Combo(["Да", "Нет"], enable_events=True, key="-FelonyAcceptedCombo-"),
     sg.Text("Дата возбуждения УД"), sg.InputText("ДД.ММ.ГГГГ", enable_events=True, key="-FelonyAcceptionDate-"),
     sg.CalendarButton("Выбрать", format="%d.%m.%Y")],
    [sg.Text("Дата заслушивания ТНО по результатам МНК"),
     sg.InputText("ДД.ММ.ГГГГ", enable_events=True, key="-TNODate-"),
     sg.CalendarButton("Выбрать", format="%d.%m.%Y", begin_at_sunday_plus=1)],
    [sg.Text("КНП окончена"), sg.Combo(["Да", "Нет"], enable_events=True, key="-InspectionIsFinishedCombo-")],
    [sg.Text("Стадии КНП"), sg.Combo(["Сбор базы", "Есть акт", "Проводятся дополнительные мероприятия", "Допы окончены",
                                      "В УФНС направлен проект решения", "Вынесено решение о привлечении",
                                      "Вынесено решение об отказе", "Отменено ВНО", "Апеляционная жалоба"],
                                     enable_events=True, key="-InspectionFinishedCombo-")],
    [sg.Text("Уплата"), sg.Combo(["УНД к оплате до акта", "Уплата после акта, но до решения", "Уплата по решению 100%",
                                  "Частичная уплата по решению", "Поступлений по решению нет"],
                                 enable_events=True, key="-PaymentCombo-")],
    [sg.Text("Причины отсутствия акта"), sg.Combo(
        ["Акт не составлялся (нет док.базы 54.1)", "ТНО не запрашивала продления, акт не составлялся",
         "ТНО установлен иной, поэтому продления не запрашивали",
         "Информация об открытой КНП у ВП в адрес УФНС не направлялась"], key="-ActFailureReasonCombo-",
        enable_events=True)],
    [sg.Text("Сумма поступившей уплаты на 5 число представления информации, тыс.руб."),
     sg.InputText(enable_events=True, key="-PaidSum-")],
    [sg.Text("Остаток (=сумма расхождений - сумма уплаты), тыс. руб"),
     sg.InputText(disabled=True, key="-ModSum-")],
    [sg.Text("Комментарии ТНО"), sg.InputText(enable_events=True, key="-CommentText-")],
    [sg.Text("Мои вопросы к ТНО"), sg.InputText(enable_events=True, key="-QuestionsText-")],
    [sg.Text("Кто пишет акт"), sg.InputText(enable_events=True, key="-AuthorText-")],
    [sg.Button("Добавить запись")],
]
window = sg.Window('TAXInput', enable_close_attempted_event=True, auto_size_text=True,
                   return_keyboard_events=True).Layout(
    [[sg.Column(layout, size=(1200, 600), scrollable=True)]])
while True:  # The Event Loop
    event, values = window.read()
    print(values)
    for value in values.items():
        try:
            if value[1] in ("", "ДД.ММ.ГГГГ"):
                # if "Date" in value[0]:
                # window["InputFrame["+(value[0])[-5:-1]].update(background_color="orange")
                # else:
                window[value[0]].update(background_color="orange")
        except TypeError:
            continue
    if (event == sg.WINDOW_CLOSE_ATTEMPTED_EVENT or event == 'Exit') and sg.popup_yes_no(
            'Do you really want to exit?') == 'Yes':
        break
    if event in ["-DeclarationNumber-", "-DeclarationPeriodCombo-", "-DeclarationYearSlip-",
                 "-ActualDeclarationCombo-"]:
        window["Добавить декларацию"].update(disabled=False)
    # print(event, values) #debug
    if event == "Добавить декларацию":
        declaration_list.append([values["-DeclarationNumber-"],
                                 "0" + str(values["-DeclarationPeriodCombo-"]) + "-" + str(
                                     values["-DeclarationYearSpin-"]),
                                 values["-ActualDeclarationCombo-"]])
        window["-DeclarationsTable-"].update(values=declaration_list)
        window["-DeclarationsQuantity-"].update("Кол-во элементов:" + str(len(declaration_list)))
        window["Добавить декларацию"].update(disabled=True)
    if event in ["-ContragentName-", "-ContragentNumber-"]:
        window["Добавить"].update(disabled=False)
    if "Date" in event:
        window["InputFrame[" + event[-5:-1]].update(background_color="red")
    if "Sum" in event:
        window[event].update(background_color="white")
        letter = values[event][-1] if values[event] else ""
        pos = False
        for comma in (",", "."):
            if comma in values[event]:
                window[event].update(value=values[event].replace(".", ","))
                pos = values[event].index(comma)
                break
        if letter not in digits and letter not in (",", "."):
            window[event].update(value=values[event][:-1], background_color="red")
        elif len(values[event]) == 1 and letter in (",", "."):
            window[event].update(value=values[event][:-1], background_color="red")
        elif letter in (",", ".") and pos != values[event].index(letter):
            window[event].update(value=values[event][:-1], background_color="red")
        elif letter in digits and not pos and len(values[event]) > 5:
            window[event].update(value=values[event][:-1])
        elif letter in digits and pos and len(values[event][pos:]) > 4:
            window[event].update(value=values[event][:-1])
        elif not "" in (values['-FraudSum-'], values['-PaidSum-']):
            window["-ModSum-"].update(value=str("{:.3f}".format(float(values['-FraudSum-'].replace(",","."))-float(values['-PaidSum-'].replace(",",".")))).replace(".",","))

    if event == "Добавить":
        contragent_list.append([values["-ContragentName-"], values["-ContragentNumber-"]])
        window["-ContragentsTable-"].update(values=contragent_list)
        window["-ContragentsQuantity-"].update("Кол-во элементов:     " + str(len(contragent_list)))
        window["Добавить"].update(disabled=True)
    if event == "-OfficeNumber-":
        window[event].update(background_color="white")
        correctString = ""
        for letter in values[event]:
            if letter not in digits:
                window[event].update(background_color="red", value=correctString)
                break
            else:
                correctString += letter
                n = correctString[:4]
                window[event].update(value=n)
    else:
        if len(values["-OfficeNumber-"]) != 0 and len(values["-OfficeNumber-"]) != 4:
            window["-OfficeNumber-"].update(background_color="red")
    if "TypeCombo-" in event:
        field = event[:-10] + "Number-"
        window[field].update(disabled=False,
                             value=(values[field][:10] if ("лицо" in values[event]) else values[field][:12]),
                             background_color="orange")
    if event in ["-CriminalNumber-", "-ContragentNumber-", "-MainContragentNumber-"]:
        window[event].update(background_color="white")
        correctString = ""
        for letter in values[event]:
            if letter not in digits:
                window[event].update(background_color="red", value=correctString)
                break
            else:
                correctString += letter
                if "ИП" in values[event[:-7] + "TypeCombo-"]:
                    n = correctString[:12]
                    window[event].update(value=n)
                    if len(n) == 12:
                        if int(n[-2]) != ((8 * int(n[-3]) + 6 * int(n[-4]) + 4 * int(n[-5]) + 9 * int(n[-6]) + 5 * int(
                                n[-7]) + 3 * int(n[-8]) + 10 * int(n[-9]) + 4 * int(n[-10]) + 2 * int(n[-11]) + 7 * int(
                            n[-12])) % 11) % 10 or int(n[-1]) != ((8 * int(n[-2]) + 6 * int(n[-3]) + 4 * int(
                            n[-4]) + 9 * int(n[-5]) + 5 * int(n[-6]) + 3 * int(n[-7]) + 10 * int(n[-8]) + 4 * int(
                            n[-9]) + 2 * int(n[-10]) + 7 * int(n[-11]) + 3 * int(n[-12])) % 11) % 10:
                            window[event].update(background_color="red")
                        else:
                            window[event].update(background_color="green")
                elif "лицо" in values[event[:-7] + "TypeCombo-"]:
                    n = correctString[:10]
                    window[event].update(value=n)
                    if len(n) == 10:
                        if int(n[-1]) != ((8 * int(n[-2]) + 6 * int(n[-3]) + 4 * int(
                                n[-4]) + 9 * int(n[-5]) + 5 * int(n[-6]) + 3 * int(n[-7]) + 10 * int(n[-8]) + 4 * int(
                            n[-9]) + 2 * int(n[-10])) % 11) % 10:
                            window[event].update(background_color="red")
                        else:
                            window[event].update(background_color="green")
    else:
        for value in ["-CriminalNumber-", "-ContragentNumber-", "-MainContragentNumber-"]:
            if len(values[value]) not in (0, 10, 12):
                window[value].update(background_color="red")
    if event == "Добавить запись":
        corrected_values.append([])
        corrected_input = []
        headers = []
        # if "" in corrected_values.values():
        # sg.popup("Ошибка",title="Ошибка")
        # else:
        for element in values.items():
            if str(element[0])[0] == "-" and str(element[0])[-1] == "-":
                corrected_input.append(list(element))
        for element in corrected_input:
            if element[0] == "-ChallengeNumber-":
                element[1] += " от " + corrected_input[corrected_input.index(element) + 1][1]
                corrected_values[-1].append(element[1])
            elif element[0] == "-ChallengeDate-":
                pass
            elif element[0] == "-DeclarationNumber-":
                pass
            elif element[0] == "-DeclarationPeriodCombo-":
                pass
            elif element[0] == "-ActualDeclarationCombo-":
                pass
            elif element[0] == "-DeclarationYearSpin-":
                pass
            elif element[0] == "-DeclarationsTable-":
                actual_declaration_str = ""
                nonactual_declaration_str = ""
                for i in range(len(declaration_list)):
                    if declaration_list[i][2] == "Да":
                        actual_declaration_str += declaration_list[i][0] + " " + declaration_list[i][1] + " (" + \
                                                  declaration_list[i][2] + ")" + ";"
                    else:
                        nonactual_declaration_str += declaration_list[i][0] + " " + declaration_list[i][1] + " (" + \
                                                     declaration_list[i][2] + ")" + ";"

                for string in (actual_declaration_str, nonactual_declaration_str):
                    element[1] = string
                    corrected_values[-1].append(element[1])
            elif element[0] == "-ContragentsTable-":
                contragent_str = ""
                for i in range(len(contragent_list)):
                    contragent_str += contragent_list[i][0] + " " + contragent_list[i][1] + ";"
                element[1] = contragent_str
                corrected_values[-1].append(element[1])
            elif element[0] == "-MainContragentName-":
                element[1] = " " + corrected_input[corrected_input.index(element) + 2][1]
                corrected_values[-1].append(element[1])
            elif element[0] == "-MainContragentNumber-":
                pass
            elif element[0] == "-Date[0,9]-":
                element[1] = " " + corrected_input[corrected_input.index(element) + 1][1]
                corrected_values[-1].append(element[1])
            elif element[0] == "-Number[0,9]-":
                pass
            elif not (str(element[1]) == "None" or ("Выбрать" in str(element[0])) or ("Type" in str(element[0])) or (
                    "-Contragent" in str(element[0]))):
                corrected_values[-1].append(element[1])
        # print(FormXLS.LoadWorkbook("aboba.xlsx"))
    if "Поиск" in event:
        corrected_input=[]
        for element in values.keys():
            if str(element)[0] == "-" and str(element)[-1] == "-":
                corrected_input.append(element)
        imported_list = FormXLS.LoadWorkbook(InputFile)
        for element in imported_list:
            if str(element[3]) == values["-CriminalNumber-"]:
                for i in range (len(corrected_input)):
                        if corrected_input[i]=="-DeclarationsTable-":
                            #declaration_list=element[i-1].split()
                            #window["Добавить декларацию"].click()
                            pass
                        elif corrected_input[i]=="-ContragentsTable-":
                            #contragent_list=str(element[i-1]).split()
                            #window["Добавить"].click()
                            pass
                        else:
                            window[corrected_input[i]].update(value=element[i] if element[i] else "Error")
    if event=="Открыть":
        InputFile=sg.popup_get_file("",no_window=True,file_types=(("Excel (*.xls, *.xlsx)", ("*.xls", "*.xlsx")),))
        window["Поиск"].update(disabled=False)
    if event=="Сохранить":
        try:
            if InputFile:
                FormXLS.SaveToWorkbook(corrected_values,InputFile)
        except NameError:
                FormXLS.SaveToWorkbook(corrected_values,sg.popup_get_file("",no_window=True,file_types=(("Excel (*.xls, *.xlsx)", ("*.xls", "*.xlsx")),),save_as=True))
    if event=="Сохранить как...":
        FormXLS.SaveToWorkbook(corrected_values,sg.popup_get_file("",no_window=True,file_types=(("Excel (*.xls, *.xlsx)", ("*.xls", "*.xlsx")),),save_as=True))