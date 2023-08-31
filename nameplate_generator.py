import os, shutil, zipfile
from docxtpl import DocxTemplate
from typing import List, Tuple, Dict, Optional


class NameplateGenerator:
    doc = DocxTemplate("template.docx")

    blocks = [
        ("01",),
        ("02", "03"),
        ("04", "05"),
        ("06", "07"),
        ("08", "09"),
        ("10",),
        ("11", "12"),
        ("13",),
        ("14",),
        ("15",),
        ("16",),
        ("17",),
        ("18",),
    ]

    floors = ["2", "3", "4", "5", "6", "7", "9"]

    # блок - одиночная комната, либо две комнаты с общим тамбуром
    # создает список таплов со всеми блоками комнат
    @staticmethod
    def generate_blocks() -> List[Tuple]:
        all_blocks = []
        for floor in NameplateGenerator.floors:
            for block in NameplateGenerator.blocks:
                all_blocks.append(tuple(f"{floor}{room}" for room in block))
        return all_blocks

    # делает из словаря с кортежами другой словарь с кортежами
    @staticmethod
    def fill_blocks(
        rooms: Dict[str, List[Tuple]]
    ) -> Dict[Tuple, List[Dict[str, List[Tuple]]]]:
        # Список с таплами блоков
        all_blocks = NameplateGenerator.generate_blocks()
        # Словарь, где ключ - блок, а значение - список студентов
        all_fill_blocks = {}
        for block in all_blocks:
            all_fill_blocks.setdefault(block, [])

        for room_number, students in rooms.items():
            for block in all_blocks:
                for block_room_number in block:
                    if room_number == block_room_number:
                        all_fill_blocks[block].append({room_number: students})
                        break
        return all_fill_blocks

    # создает таблички всех комнат
    def generate_nameplates(rooms: Dict[str, List[Tuple]]) -> None:
        all_fill_blocks = NameplateGenerator.fill_blocks(rooms)

        # создаем директорию с табличками
        if os.path.exists(os.path.join(os.path.abspath("."), "nameplates")):
            shutil.rmtree(os.path.join(os.path.abspath("."), "nameplates"))
        os.mkdir(os.path.join(os.path.abspath("."), "nameplates"))

        for block, rooms in all_fill_blocks.items():
            if rooms != []:
                NameplateGenerator.nameplate_generate(
                    rooms
                )

        # Кринжовое создание архива с табличками и удаление директории
        z = zipfile.ZipFile("nameplates.zip", "w")
        for file in os.listdir("./nameplates"):
            z.write(f"./nameplates/{file}")
        z.close()
        shutil.rmtree(os.path.join(os.path.abspath("."), "nameplates"))

    # создает одну табличку
    @staticmethod
    def nameplate_generate(
        rooms: List[Dict[str, List[Tuple]]]
    ) -> None:
        
        has_neighbor = len(rooms) == 2
        first_room_number = list(rooms[0].keys())[0]
        second_room_number = list(rooms[1].keys())[0] if has_neighbor else None

        rooms_tables = {first_room_number: [], second_room_number: []}
        for room in rooms:
            for room_number, students in room.items():
                for student in students:
                    rooms_tables[room_number].append({"surname": student[0],
                                            "name": student[1],
                                            "patronymic": student[2],
                                            "course": student[3]})

        context = {
            "has_neighbor": has_neighbor,
            "first_room_number": first_room_number,
            "second_room_number": second_room_number,
            "first_room": rooms_tables[first_room_number],
            "second_room": rooms_tables[second_room_number],
        }

        filename = str(first_room_number)
        if has_neighbor:
            filename += f"_{second_room_number}"
        NameplateGenerator.doc.render(context)
        NameplateGenerator.doc.save(f"./nameplates/{filename}.docx")
