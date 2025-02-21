from dataclasses import dataclass
import json
import re

re_xref_id = re.compile(r"<REF (.*?)>")
re_xref_em_id = re.compile(r"<EMID (.*?)>")


@dataclass
class DBEntry:
    name: str
    guid: str
    belongs_to: str
    contents: dict[str, str]

    def get_text(self, lang_id: int = 13) -> str:
        return self.contents.get(str(lang_id))


class TextDB:
    def __init__(self, entries: dict[str, DBEntry]):
        self.db_entries = entries
        self.index_name = {}
        self.default_lang = 13
        self._create_index_name()
        self._filter_data()

    def _create_index_name(self):
        for guid, entry in self.db_entries.items():
            self.index_name[entry.name] = guid

    def _filter_data(self):
        for entry in self.db_entries.values():
            for lang_id, text in entry.contents.items():
                entry.contents[lang_id] = text.replace("\n", "").replace("\r", "")

    def set_default_lang(self, lang_id: int):
        self.default_lang = lang_id

    def get_entry_by_name(self, name: str) -> DBEntry:
        guid = self.index_name.get(name)
        if guid is None:
            return None
        return self.db_entries.get(guid)

    def get_entry_by_guid(self, guid: str) -> DBEntry:
        return self.db_entries.get(guid)

    def get_text_by_guid(self, guid: str, lang_id: int = None) -> str:
        if lang_id is None:
            lang_id = self.default_lang
        entry = self.db_entries.get(guid)
        if entry is None:
            return None
        return entry.contents.get(str(lang_id))

    def get_text_by_name(self, name: str, lang_id: int = None) -> str:
        entry = self.get_entry_by_name(name)
        if entry is None:
            return None
        return self.get_text_by_guid(entry.guid, lang_id)


def load_text_db(db_json_path: str) -> TextDB:
    db_json = []
    with open(db_json_path, "r", encoding="utf-8") as f:
        db_json = json.load(f)

    db_entries = {}
    for i, entry in enumerate(db_json):
        if not entry:
            continue
        db_entry = DBEntry(
            name=entry["name"],
            guid=entry["guid"],
            belongs_to=entry["belongs_to"],
            contents=entry["contents"],
        )
        db_entries[db_entry.guid] = db_entry

    return TextDB(db_entries)


if __name__ == "__main__":
    text_db = load_text_db("texts_db.json")

    assert text_db.get_entry_by_guid("38e2f690-384c-4db5-824e-a289d2335b53") is not None
    assert text_db.get_entry_by_name("NpcName_NN_86") is not None
    assert text_db.get_text_by_name("PanelTutorial_TITLE_61", lang_id=13) is not None
    print(text_db.get_text_by_guid("38e2f690-384c-4db5-824e-a289d2335b53", lang_id=13))
