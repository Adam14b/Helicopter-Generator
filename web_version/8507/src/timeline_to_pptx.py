from streamlit_pptx import *
from datetime import date, timedelta
from enum import Enum

class Event:
    def __init__(self, event_type: str = '-', event_info: list = []) -> None:
        self.event_type: str = event_type
        self.event_info: list = event_info

    def __repr__(self) -> str:
        return f'Event({self.event_type}, {self.event_info})'

    def to_dict(self):
        event_info_serialized = []
        for item in self.event_info:
            if isinstance(item, date):
                event_info_serialized.append(item.strftime('%Y-%m-%d'))
            else:
                event_info_serialized.append(item)
        return {
            'event_type': self.event_type,
            'event_info': event_info_serialized,
        }

    @classmethod
    def from_dict(cls, data: dict):
        event_info_deserialized = []
        for item in data['event_info']:
            if isinstance(item, str) and '-' in item:
                event_info_deserialized.append(date.fromisoformat(item))
            else:
                event_info_deserialized.append(item)
        return cls(
            event_type=data['event_type'],
            event_info=event_info_deserialized,
        )

class Task:
    def __init__(self, task: str = '', tags: str = '', effect: str = '', int_: str = '', complexity: str = '', chd: str = '', events: list = [], start: date = date.today(), finish: date = date.today()) -> None:
        self.task: str = task
        self.tags: str = tags
        self.effect: str = effect
        self.int_: str = int_
        self.complexity: str = complexity
        self.chd: str = chd
        self.events: list[Event] = events
        self.start: date = start
        self.finish: date = finish

    def __repr__(self) -> str:
        events = "\n    ".join([''] + [str(event) for event in self.events])
        return f'Task({self.task}, {self.tags}, {self.start}, {self.finish}, {events}\n)'

    def __len__(self):
        return len(self.events)

    def to_dict(self):
        return {
            'task': self.task,
            'tags': self.tags,
            'effect': self.effect,
            'int_': self.int_,
            'complexity': self.complexity,
            'chd': self.chd,
            'events': [event.to_dict() for event in self.events],
            'start': self.start.strftime('%Y-%m-%d'),
            'finish': self.finish.strftime('%Y-%m-%d')
        }

    @classmethod
    def from_dict(cls, data: dict):
        events = [Event.from_dict(event_data) for event_data in data['events']]
        start = date.fromisoformat(data['start']) if isinstance(data['start'], str) else data['start']
        finish = date.fromisoformat(data['finish']) if isinstance(data['finish'], str) else data['finish']
        return cls(
            task=data['task'],
            tags=data['tags'],
            effect=data['effect'],
            int_=data['int_'],
            complexity=data['complexity'],
            chd=data['chd'],
            events=events,
            start=start,
            finish=finish,
        )

colors = {
    'серый': GRAY,
    'зелёный': GREEN,
    'жёлтый': YELLOW,
    'красный': RED,
}

class EventType(Enum):
    SUCCESS = 'успех [зелёный флаг]'
    PLAN = 'план [серый флаг]'
    MOVE_RISK = 'риск переноса [жёлтый флаг]'
    FAIL = 'провал [красный флаг]'
    MOVE = 'перенос [красный флаг ---> серый флаг]'
    FINALLY = 'сдача с опозданием [красный флаг ---> жёлтый флаг]'
    SPEEDRUN = 'сдача раньше срока [зелёный флаг <--- зелёный флаг]'
    IFT = 'ИФТ [гаечный ключ]'
    NT = 'НТ [шестерёнка]'
    PSI = 'ПСИ [коробка со стрелкой]'
    PROD = 'Прод [ракета наклонена]'
    MVP = 'MVP [пимпочка]'
    PILOT = 'Пилот [ракета на старте]'

def process(day: date, tasks: list[Task], file_name: str):
    tasks_per_quarter = {}
    total_tasks = 0
    dates = []

    for task in tasks:
        start_date = task.start
        finish_date = task.finish
        dates.extend([start_date, finish_date])

        for date_ in [start_date, finish_date]:
            quarter = (date_.year, (date_.month - 1) // 3 + 1)
            tasks_per_quarter.setdefault(quarter, 0)
            tasks_per_quarter[quarter] += 1
            total_tasks += 1
        for event in task.events:
            if event.event_type != '-':
                if event.event_type in [EventType.MOVE.value, EventType.FINALLY.value, EventType.SPEEDRUN.value]:
                    dates.extend([event.event_info[0], event.event_info[1]])
                    for date_ in [event.event_info[0], event.event_info[1]]:
                        quarter = (date_.year, (date_.month - 1) // 3 + 1)
                        tasks_per_quarter.setdefault(quarter, 0)
                        tasks_per_quarter[quarter] += 1
                        total_tasks += 1
                else:
                    date_ = event.event_info[0]
                    dates.append(date_)
                    quarter = (date_.year, (date_.month - 1) // 3 + 1)
                    tasks_per_quarter.setdefault(quarter, 0)
                    tasks_per_quarter[quarter] += 1
                    total_tasks += 1

    earliest_date = min(dates)
    latest_date = max(dates)

    quarters_with_tasks = sorted(tasks_per_quarter.keys())

    report = Report(day, quarters_with_tasks, tasks_per_quarter, total_tasks)

    running_top = 90
    task_height = 420 // len(tasks) if len(tasks) > 0 else 50

    for task in tasks:
        timeline = Timeline(report, y_top=running_top, y_bottom=running_top + task_height, start_date=task.start, final_date=task.finish, description=[task.task, task.tags, task.effect, task.int_, task.complexity, task.chd])
        for event in task.events:
            if event.event_type == EventType.SUCCESS.value:
                timeline.add_pictogram(base_flag, GREEN, event.event_info[0], note=event.event_info[1], fill=day > event.event_info[0])

            elif event.event_type == EventType.PLAN.value:
                timeline.add_pictogram(base_flag, GRAY, event.event_info[0], note=event.event_info[1], fill=day > event.event_info[0])

            elif event.event_type == EventType.MOVE_RISK.value:
                timeline.add_pictogram(base_flag, YELLOW, event.event_info[0], note=event.event_info[1], fill=day > event.event_info[0])

            elif event.event_type == EventType.FAIL.value:
                timeline.add_pictogram(base_flag, RED, event.event_info[0], note=event.event_info[1], fill=day > event.event_info[0])

            elif event.event_type == EventType.MOVE.value:
                timeline.add_pictogram(base_flag, RED, event.event_info[0], fill=day > event.event_info[0])
                timeline.add_pictogram(base_flag, GRAY, event.event_info[1], fill=day > event.event_info[0])
                timeline.add_arrow(event.event_info[0], event.event_info[1])
                timeline.add_comment(event.event_info[2], event.event_info[1], size=120)

            elif event.event_type == EventType.FINALLY.value:
                timeline.add_pictogram(base_flag, RED, event.event_info[0], fill=day > event.event_info[0])
                timeline.add_pictogram(base_flag, YELLOW, event.event_info[1], fill=day > event.event_info[0])
                timeline.add_arrow(event.event_info[0], event.event_info[1])
                timeline.add_comment(event.event_info[2], event.event_info[1], size=120)

            elif event.event_type == EventType.SPEEDRUN.value:
                timeline.add_pictogram(base_flag, GREEN, event.event_info[0], fill=day > event.event_info[0])
                timeline.add_pictogram(base_flag, GREEN, event.event_info[1], fill=day > event.event_info[0])
                timeline.add_arrow(event.event_info[1], event.event_info[0])

            elif event.event_type == EventType.IFT.value:
                timeline.add_pictogram(key_fig, colors[event.event_info[2]], event.event_info[0], note=event.event_info[1])

            elif event.event_type == EventType.NT.value:
                timeline.add_pictogram(gear_fig, colors[event.event_info[2]], event.event_info[0], note=event.event_info[1])

            elif event.event_type == EventType.PSI.value:
                timeline.add_pictogram(box_fig, colors[event.event_info[2]], event.event_info[0], note=event.event_info[1])

            elif event.event_type == EventType.PROD.value:
                timeline.add_pictogram(rocket_fig, colors[event.event_info[2]], event.event_info[0], note=event.event_info[1])

            elif event.event_type == EventType.MVP.value:
                timeline.add_pictogram(pin_fig, colors[event.event_info[2]], event.event_info[0], note=event.event_info[1])

            elif event.event_type == EventType.PILOT.value:
                timeline.add_pictogram(pilot_fig, colors[event.event_info[2]], event.event_info[0], note=event.event_info[1])

        running_top += task_height

    report.save(f'../contents/{file_name}')
