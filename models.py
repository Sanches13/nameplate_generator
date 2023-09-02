from pydantic import BaseModel

class Student(BaseModel):
    surname: str
    name: str
    patronymic: str
    course: str

Room = dict[int, list[Student]]

class NameplateConfig(BaseModel):
    block_config: list[tuple[int, ...]]
    rooms: Room