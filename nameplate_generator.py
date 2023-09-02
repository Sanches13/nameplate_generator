import os, zipfile, tempfile, copy
from docxtpl import DocxTemplate
from typing import List, Tuple, Dict
from models import Student, Room, NameplateConfig
import concurrent.futures


class NameplateGenerator:
    __doc = DocxTemplate("template.docx")

    @staticmethod
    def __get_neighbour(
        current_rooms_config: list[Room],
        blocks_config: list[tuple[int]],
        room_number: int,
    ) -> int | None:
        floor = room_number // 100
        main_part = room_number % 100

        for block in blocks_config:
            if len(block) != 1 and main_part in block:
                block_as_list = list(block)
                block_as_list.remove(main_part)
                neighbour_main_part = block_as_list[0]
                for neighbour_room_number, neighbours in current_rooms_config.items():
                    if floor == neighbour_room_number // 100 and neighbour_main_part == neighbour_room_number % 100:
                        return neighbour_room_number

    @staticmethod
    def __generate_blocks(config: NameplateConfig) -> list[list[Room]]:
        blocks_list = []
        current_rooms_config = copy.deepcopy(config.rooms)
        for room_number, students in config.rooms.items():
            if room_number not in current_rooms_config.keys():
                continue
            block_rooms = [{room_number: students}]
            del current_rooms_config[room_number]
            neighbour = NameplateGenerator.__get_neighbour(
                current_rooms_config, config.block_config, room_number
            )
            if neighbour:
                block_rooms += [{neighbour: config.rooms[neighbour]}]
                del current_rooms_config[neighbour]
            blocks_list.append(block_rooms)
        return blocks_list

    # создает таблички всех комнат
    def generate_nameplates(config: NameplateConfig) -> None:
        with tempfile.TemporaryDirectory() as tmpdir, zipfile.ZipFile(
            "nameplates.zip", "w"
        ) as z:
            with concurrent.futures.ThreadPoolExecutor() as executor:
                blocks = NameplateGenerator.__generate_blocks(config)
                for rooms in blocks:
                    future = executor.submit(
                        NameplateGenerator.__nameplate_generate, tmpdir, rooms
                    )
                    future.result()
                    
            for file in os.listdir(tmpdir):
                z.write(os.path.join(tmpdir, file), f"./nameplates/{file}")

    # создает одну табличку
    @staticmethod
    def __nameplate_generate(dirname: str, rooms: list[Room]) -> None:
        has_neighbor = len(rooms) == 2
        first_room_number = list(rooms[0].keys())[0]
        second_room_number = list(rooms[1].keys())[0] if has_neighbor else None

        rooms_tables = {first_room_number: [], second_room_number: []}
        for room in rooms:
            for room_number, students in room.items():
                for student in students:
                    rooms_tables[room_number].append(
                        {
                            "surname": student.surname,
                            "name": student.name,
                            "patronymic": student.patronymic,
                            "course": student.course
                        }
                    )

        context = {
            "has_neighbor": has_neighbor,
            "first_room_number": first_room_number,
            "second_room_number": second_room_number,
            "first_room": rooms_tables[first_room_number],
            "second_room": rooms_tables[second_room_number]
        }

        filename = str(first_room_number)
        if has_neighbor:
            filename += f"_{second_room_number}"
        NameplateGenerator.__doc.render(context)
        NameplateGenerator.__doc.save(os.path.join(dirname, f"{filename}.docx"))
