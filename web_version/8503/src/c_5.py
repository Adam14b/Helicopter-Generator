import streamlit as st
from datetime import date
from time import time
from timeline_to_pptx import Task, Event, process, EventType, colors
from io import BytesIO
from functools import partial
from random import randint
import json

def custom_serializer(obj):
    if isinstance(obj, Task):
        return obj.to_dict()
    elif isinstance(obj, date):
        return obj.strftime('%Y-%m-%d')
    raise TypeError(f"Type {type(obj)} not serializable")

st.title('Сервис автогенерации хеликоптеров')
st.subheader('Нестандартные задачи')

with open('../contents/example_for_pptx.streamlit', 'rb') as rfile:
    st.download_button(
        label="Скачать пример streamlit",
        data=BytesIO(rfile.read()),
        file_name="example.streamlit",
    )

with open('../contents/example_generated.pptx', 'rb') as rfile:
    st.download_button(
        label="Скачать пример pptx",
        data=BytesIO(rfile.read()),
        file_name="example.pptx",
    )

refresh = st.button('Сбросить состояние (вся текущая сессия будет удалена)')
if refresh:
    st.session_state.timelines = [Task()]

uploaded_file = st.file_uploader('Импортировать')
day = None
if uploaded_file is not None:
    main = json.load(uploaded_file)
    st.session_state.timelines = [Task.from_dict(task) for task in main['timelines']]
    st.session_state.day = date.fromisoformat(main['day'])

loaded_from_file = st.button('Загрузить предыдущее состояние')
if loaded_from_file:
    with open('../contents/save.streamlit', 'r') as rfile:
        main = json.load(rfile)
        st.session_state.timelines = [Task.from_dict(task) for task in main['timelines']]
        st.session_state.day = date.fromisoformat(main['day'])

if 'timelines' not in st.session_state:
    st.session_state.timelines: list[Task] = [Task()]
if 'day' not in st.session_state:
    st.session_state.day: date = date.today()
st.session_state.day = st.date_input('Дата составления отчёта:', value=st.session_state.day)

today = date.today()

def require_date(label: str = 'Дата:', default_value: date = today, key: int = 0):
    default_value = default_value if default_value is not None else today
    return st.date_input(label, default_value, key=key)

def require_comment(label: str = 'Комментарий:', default_value: str = '', key: int = 0):
    default_value = default_value if default_value is not None else  ''
    return st.text_input(label, default_value, key=key)

color_options = list(colors.keys())
color_choices = {v: k for k, v in enumerate(color_options)}

def require_color(label: str = 'Цвет:', default_value: str = 'серый', key: int = 0):
    index = color_choices.get(default_value, 0)
    return st.selectbox(label, color_options, index, key=key)

events = {
    '-': [],
    EventType.SUCCESS.value: [require_date, require_comment],
    EventType.PLAN.value: [require_date, require_comment],
    EventType.MOVE_RISK.value: [require_date, require_comment],
    EventType.FAIL.value: [require_date, require_comment],
    EventType.MOVE.value: [partial(require_date, 'Дата откуда:'), partial(require_date, 'Дата куда:'), require_comment],
    EventType.FINALLY.value: [partial(require_date, 'Дата откуда:'), partial(require_date, 'Дата куда:'), require_comment],
    EventType.SPEEDRUN.value: [partial(require_date, 'Дата откуда:'), partial(require_date, 'Дата куда:')],
    EventType.IFT.value: [require_date, require_comment, require_color],
    EventType.NT.value: [require_date, require_comment, require_color],
    EventType.PSI.value: [require_date, require_comment, require_color],
    EventType.PROD.value: [require_date, require_comment, require_color],
    EventType.MVP.value: [require_date, require_comment, require_color],
    EventType.PILOT.value: [require_date, require_comment, require_color],
}

event_options = list(events.keys())
event_param_count = {k: len(v) for k, v in events.items()}

unique_key = 0

def show_event(timeline_i: int, event_i: int):
    global unique_key
    unique_key += 1
    if st.button(f'Удалить событие №{timeline_i + 1}.{event_i + 1}', key=f"del_event_{timeline_i}_{event_i}"):
        st.session_state.timelines[timeline_i].events.pop(event_i)
        st.experimental_rerun()

    default_index = event_options.index(st.session_state.timelines[timeline_i].events[event_i].event_type)
    unique_key += 1
    event_type = st.selectbox(f"Задача №{timeline_i + 1}. Событие №{event_i + 1}", options=event_options, index=default_index, key=f"event_type_{timeline_i}_{event_i}")

    st.session_state.timelines[timeline_i].events[event_i].event_type = event_type

    if len(st.session_state.timelines[timeline_i].events[event_i].event_info) != event_param_count[event_type]:
        st.session_state.timelines[timeline_i].events[event_i].event_info = [None] * event_param_count[event_type]

    for k, input_field in enumerate(events[event_type]):
        unique_key += 1
        default_value = st.session_state.timelines[timeline_i].events[event_i].event_info[k]
        st.session_state.timelines[timeline_i].events[event_i].event_info[k] = input_field(key=f"event_{timeline_i}_{event_i}_{k}", default_value=default_value)

def show_timeline(timeline_i: int):
    global unique_key
    unique_key += 1
    if st.button(f'Удалить задачу №{timeline_i + 1}', key=f"del_task_{timeline_i}"):
        st.session_state.timelines.pop(timeline_i)
        st.experimental_rerun()
    else:
        with st.expander(f'Задача №{timeline_i + 1}'):
            unique_key += 1
            st.session_state.timelines[timeline_i].task = st.text_input(f'Задача №{timeline_i + 1}. Описание задачи', st.session_state.timelines[timeline_i].task, key=f"task_{timeline_i}", placeholder='Описание задачи')
            unique_key += 1
            st.session_state.timelines[timeline_i].tags = st.text_input(f'Задача №{timeline_i + 1}. Тэги', st.session_state.timelines[timeline_i].tags, key=f"tags_{timeline_i}", placeholder='Тэги')
            unique_key += 1
            st.session_state.timelines[timeline_i].effect = st.text_input(f'Задача №{timeline_i + 1}. Эффект', st.session_state.timelines[timeline_i].effect, key=f"effect_{timeline_i}", placeholder='Эффект')
            unique_key += 1
            st.session_state.timelines[timeline_i].int_ = st.text_input(f'Задача №{timeline_i + 1}. Интеграции', st.session_state.timelines[timeline_i].int_, key=f"int_{timeline_i}", placeholder='Интеграции')
            unique_key += 1
            st.session_state.timelines[timeline_i].complexity = st.text_input(f'Задача №{timeline_i + 1}. Сложность', st.session_state.timelines[timeline_i].complexity, key=f"complexity_{timeline_i}", placeholder='Сложность')
            unique_key += 1
            st.session_state.timelines[timeline_i].chd = st.text_input(f'Задача №{timeline_i + 1}. Человекодни', st.session_state.timelines[timeline_i].chd, key=f"chd_{timeline_i}", placeholder='Человекодни')
            unique_key += 1
            st.session_state.timelines[timeline_i].start = st.date_input(f'Задача №{timeline_i + 1}. Начало', st.session_state.timelines[timeline_i].start, key=f"start_{timeline_i}")
            unique_key += 1
            st.session_state.timelines[timeline_i].finish = st.date_input(f'Задача №{timeline_i + 1}. Конец', st.session_state.timelines[timeline_i].finish, key=f"finish_{timeline_i}")

            for event_i in range(len(st.session_state.timelines[timeline_i])):
                show_event(timeline_i, event_i)

            unique_key += 1
            if st.button("Добавить событие", key=f"add_event_{timeline_i}"):
                st.session_state.timelines[timeline_i].events.append(Event())
                show_event(timeline_i, len(st.session_state.timelines[timeline_i]) - 1)

for timeline_i in range(len(st.session_state.timelines)):
    if st.session_state.timelines[timeline_i] is not None:
        show_timeline(timeline_i)

if st.button("Добавить задачу"):
    st.session_state.timelines.append(Task())
    show_timeline(len(st.session_state.timelines) - 1)

if st.button('Сохранить текущее состояние'):
    with open('..\contents\save.streamlit', 'w') as wfile:
        json.dump({'timelines': st.session_state.timelines, 'day': st.session_state.day}, wfile, ensure_ascii=False, default=custom_serializer, indent=4)

if st.button('Обработать'):
    tm = int(time())
    filename = f'../contents/helicopter_{tm}.pptx'
    process(st.session_state.day, st.session_state.timelines, filename)
    with open(filename, 'rb') as rfile:
        st.download_button('Скачать pptx', data=BytesIO(rfile.read()), file_name=filename)

    with open('../contents/save_tmp.streamlit', 'w') as wfile:
        json.dump({'timelines': st.session_state.timelines, 'day': st.session_state.day}, wfile, ensure_ascii=False, default=custom_serializer, indent=4)
    with open('../contents/save_tmp.streamlit', 'rb') as rfile:
        st.download_button('Экспортировать текущее состояние', data=BytesIO(rfile.read()), file_name=f'../contents/helicopter_{tm}.streamlit')

st.write("Внутренняя структура данных:", st.session_state.timelines)
