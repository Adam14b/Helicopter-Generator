import pandas as pd
from report import Report
from timeline import Timeline
from datetime import date
from config import fig_dict

def _str_to_date(date_str):
    try:
        year, month, day = tuple(map(int, str(date_str).split(' ')[0].split('-')))
    except Exception:
        day, month, year = tuple(map(int, str(date_str).split(' ')[0].split('.')))
    return date(year, month, day)

def process(in_file: str, out_file: str):
    data = pd.read_excel(in_file)
    dates = []
    tasks_per_quarter = {}

    total_tasks = 0  # Общее количество задач

    for i, (main, param_1, param_2) in enumerate(zip(data.columns[::3], data.columns[1::3], data.columns[2::3])):
        main_column = data[main]
        param_1_column = data[param_1]
        param_2_column = data[param_2]

        start_date = _str_to_date(param_1_column[1])
        final_date = _str_to_date(param_2_column[1])
        dates.extend([start_date, final_date])

        # Добавляем задачи в кварталы
        for date_ in [start_date, final_date]:
            quarter = (date_.year, (date_.month - 1) // 3 + 1)
            tasks_per_quarter.setdefault(quarter, 0)
            tasks_per_quarter[quarter] += 1
            total_tasks += 1

        for k in range(2, len(main_column)):
            if main_column[k] == 'комментарий':
                date_ = _str_to_date(param_1_column[k])
                dates.append(date_)
            elif main_column[k] == 'перенос':
                date1 = _str_to_date(param_1_column[k])
                date2 = _str_to_date(param_2_column[k])
                dates.extend([date1, date2])
            elif main_column[k] in fig_dict.keys() or isinstance(main_column[k], str) and '/' in main_column[k] and main_column[k].split('/')[0] in fig_dict.keys():
                date_ = _str_to_date(param_1_column[k])
                dates.append(date_)
                quarter = (date_.year, (date_.month - 1) // 3 + 1)
                tasks_per_quarter.setdefault(quarter, 0)
                tasks_per_quarter[quarter] += 1
                total_tasks += 1

    earliest_date = min(dates)
    latest_date = max(dates)

    # Создаем отсортированный список кварталов с задачами
    quarters_with_tasks = sorted(tasks_per_quarter.keys())

    report = Report(_str_to_date(data.columns[0]), quarters_with_tasks, tasks_per_quarter, total_tasks)

    running_top = 90
    for i, (main, param_1, param_2) in enumerate(zip(data.columns[::3], data.columns[1::3], data.columns[2::3])):
        main_column = data[main]
        param_1_column = data[param_1]
        param_2_column = data[param_2]

        n = len(main_column)
        description = main_column[1].split('\n')
        start_date = _str_to_date(param_1_column[1])
        final_date = _str_to_date(param_2_column[1])

        timeline_height = int(param_2_column[0])

        timeline = Timeline(report, running_top, running_top + timeline_height, start_date, final_date, description, show_start_date=True if pd.isna(param_1_column[0]) else False)

        golds = []
        for k in range(2, n):
            if main_column[k] == 'комментарий':
                timeline.add_comment(param_2_column[k], _str_to_date(param_1_column[k]), 120)
            elif main_column[k] == 'перенос':
                timeline.add_arrow(_str_to_date(param_1_column[k]), _str_to_date(param_2_column[k]))
            elif main_column[k] == 'слиток':
                golds.append(str(param_1_column[k]))
            elif main_column[k] in fig_dict.keys() or (isinstance(main_column[k], str) and '/' in main_column[k] and main_column[k].split('/')[0] in fig_dict.keys()):
                write_date = _str_to_date(param_1_column[k]) != final_date and (param_2_column[k] != '--')
                note = str(param_2_column[k]) if not pd.isna(param_2_column[k]) and (param_2_column[k] != '--') else ''
                timeline.add_figure(main_column[k], _str_to_date(param_1_column[k]), write_date, note)

        for k, gold in enumerate(golds):
            timeline.add_gold(gold, timeline_height * (k / len(golds)))

        running_top += timeline_height

    report.save(f'../contents/{out_file}')