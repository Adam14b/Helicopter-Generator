import pandas as pd
from report import Report
from timeline import Timeline
from datetime import date
from config import fig_dict
def _str_to_date(date_str):
    try:
        year, month, day = tuple( map(int, str(date_str).split(' ')[0].split('-')) )
    except Exception:
        day, month, year = tuple( map(int, str(date_str).split(' ')[0].split('.')) )
    return date(year, month, day)
def process(in_file: str, out_file: str):
    data = pd.read_excel(in_file)

    report = Report(_str_to_date(data.columns[0]))

    running_top = 90
    for i, (main, param_1, param_2) in enumerate(zip(data.columns[::3], data.columns[1::3], data.columns[2::3])):
        main_column = data[main]
        param_1_column = data[param_1]
        param_2_column = data[param_2]
        #print(i, main_column, param_1_column, param_2_column) 
        n = len(main_column)
        print(main_column[1])
        description = main_column[1].split('\n')
        start_date = _str_to_date(param_1_column[1])
        final_date = _str_to_date(param_2_column[1])

        timeline_height = int(param_2_column[0])
        print(running_top)
        print(timeline_height)

        timeline = Timeline(report, running_top, running_top + timeline_height, start_date, final_date, description, show_start_date=True if pd.isna(param_1_column[0]) else False)

        golds = []
        for k in range(2, n):
            if main_column[k] == 'комментарий':
                timeline.add_comment(param_2_column[k], _str_to_date(param_1_column[k]), 120)
            elif main_column[k] == 'перенос':
                timeline.add_arrow(_str_to_date(param_1_column[k]), _str_to_date(param_2_column[k]))
            elif main_column[k] == 'слиток':
                golds.append(str(param_1_column[k]))
            elif main_column[k] in fig_dict.keys() or isinstance(main_column[k], str) and main_column[k].find('/') and main_column[k].split('/')[0] in fig_dict.keys():
                write_date = _str_to_date(param_1_column[k]) != final_date and (param_2_column[k] != '--')
                note = str(param_2_column[k]) if not pd.isna(param_2_column[k]) and (param_2_column[k] != '--') else ''
                timeline.add_figure(main_column[k], _str_to_date(param_1_column[k]), write_date, note)

        for k, gold in enumerate(golds):
            timeline.add_gold(gold, timeline_height*(k / len(golds)))

        running_top += timeline_height

    report.save(out_file)
