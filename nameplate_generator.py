import os, shutil, zipfile
from docxtpl import DocxTemplate
from typing import List, Tuple, Dict, Optional

class NameplateGenerator:
    
    doc = DocxTemplate("template.docx")

    blocks = [("01",),
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
             ("18",)]

    floors = ["2", "3", "4", "5", "6", "7", "9"]

    # делает из массива с кортежами словарь с кортежами
    @staticmethod
    def fill_rooms(students: List[Tuple]) -> Dict[int, List[Tuple]]:
        rooms = {}
        for student in students:
            rooms.setdefault(student[0], [])
            rooms[student[0]].append(student)
        return rooms
    
    # создает таплы со всеми блоками комнат
    @staticmethod
    def generate_blocks() -> List[Tuple]:
        all_blocks = []
        for floor in NameplateGenerator.floors:
            for block in NameplateGenerator.blocks:
                all_blocks.append(tuple(f"{floor}{room}" for room in block))
        return all_blocks

    # делает из словаря с кортежами другой словарь с кортежами
    @staticmethod
    def fill_blocks(students: List[Tuple]) -> Dict[Tuple, Dict[Tuple, List[Tuple]]]:
        rooms = NameplateGenerator.fill_rooms(students)
        all_blocks = NameplateGenerator.generate_blocks()
        all_fill_blocks = {}
        for block in all_blocks:
            all_fill_blocks.setdefault(block, [])

        for room_number in rooms.keys():
            for block in all_blocks:
                for block_room_number in block:
                    if str(room_number) == block_room_number:
                        all_fill_blocks[block].append(rooms[room_number])
                        break
        return all_fill_blocks

    # создает таблички всех комнат
    def generate_nameplates(students: List[Tuple]) -> None:
        all_fill_blocks = NameplateGenerator.fill_blocks(students)

        # создаем директорию с табличками
        if os.path.exists(os.path.join(os.path.abspath("."), "nameplates")):
            shutil.rmtree(os.path.join(os.path.abspath("."), "nameplates"))
        os.mkdir(os.path.join(os.path.abspath("."), "nameplates"))

        for block in all_fill_blocks.keys():
            first_room_number = int(block[0])
            second_room_number = int(block[1]) if len(block) == 2 else None
            NameplateGenerator.nameplate_generate(all_fill_blocks[block],
                                                  first_room_number,
                                                  second_room_number)
        
        # Кринжовое создание архива с табличками и удаление директории
        z = zipfile.ZipFile('nameplates.zip', 'w')
        for file in os.listdir("./nameplates"):
            z.write(f"./nameplates/{file}")
        z.close()
        shutil.rmtree(os.path.join(os.path.abspath("."), "nameplates"))

    # создает одну табличку
    @staticmethod
    def nameplate_generate(blocks: List[Tuple],
                           first_room_number: int,
                           second_room_number: Optional[int]) -> None:
        
        rooms = {first_room_number: [], second_room_number: []}
        for room in blocks:
            for student in room:
                rooms[student[0]].append({"surname": student[1],
                                        "name": student[2],
                                        "patronymic": student[3],
                                        "course": student[4]})
        context = {
            "has_neighbor": second_room_number,
            "first_room_number": first_room_number,
            "second_room_number": second_room_number,
            "first_room": rooms[first_room_number],
            "second_room": rooms[second_room_number]
        }

        filename = str(first_room_number)
        if second_room_number:
            filename += f"_{second_room_number}"
        NameplateGenerator.doc.render(context)
        NameplateGenerator.doc.save(f"./nameplates/{filename}.docx")