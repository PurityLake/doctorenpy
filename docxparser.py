import docx
from dataclasses import dataclass


class Character:
    __slots__ = ("name", "varname")

    def __init__(self, name: str) -> None:
        self.name = name
        self.varname = name.lower().strip().replace(" ", "_")

    def __hash__(self) -> int:
        return hash(self.name + self.varname)

    def __eq__(self, value: object) -> bool:
        if type(value) is Character:
            return self.name == value.name and self.varname == value.varname
        return False


if __name__ == "__main__":
    lines = []
    characters = set()

    doc = docx.Document("example.docx")
    character: Character | None = None

    default_font_size = -1

    for para in doc.paragraphs:
        if para.text.startswith("("):
            lines.append(f"    # {para.text}\n\n")
            continue
        line = ""
        for run in para.runs:
            if default_font_size == -1 and run.font.size is not None:
                default_font_size = run.font.size.pt
            if run.bold and run.text.strip().endswith(":"):
                character = Character(run.text.strip()[:-1])
                characters.add(character)
                continue
            temp = run.text
            if run.bold:
                temp = "{b}" + temp + "{/b}"
            if run.italic:
                temp = "{i}" + temp + "{/i}"
            if run.underline:
                temp = "{u}" + temp + "{/u}"
            if run.font.size is not None:
                print(run.font.size.pt)
                if run.font.size.pt > default_font_size:
                    temp = (
                        "{size=+"
                        + str(int(run.font.size.pt - default_font_size))
                        + "}"
                        + temp
                        + "{/size}"
                    )
                if run.font.size.pt < default_font_size:
                    temp = (
                        "{size=-"
                        + str(int(default_font_size - run.font.size.pt))
                        + "}"
                        + temp
                        + "{/size}"
                    )
            line += temp
        if len(line) > 0:
            line = line.replace("\u2018", "'").replace("\u2019", "'")
            line = line.replace("\u201C", '\\"').replace("\u201D", '\\"')
            line = line.replace("\u2026", "...")

            if character is not None:
                lines.append(f'    {character.varname} "{line}"\n\n')
                character = None
            else:
                lines.append(f'    "{line}"\n\n')

    with open("output.rpy", "w") as f:
        f.write("init:\n")

        for character in characters:
            f.write(f'    $ {character.varname} = Character("{character.name}")\n')

        f.write("\nlabel start:\n")

        for line in lines:
            f.write(line)
