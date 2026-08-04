import json
import os


class SettingsManager:

    def __init__(self):

        self.filename = os.path.join(
            "config",
            "settings.json",
        )

        self.settings = {
            "workbook": "",
            "pdf_library": "",
            "output_folder": "",
        }

        self.load()

    ##################################################

    def load(self):

        if not os.path.exists(self.filename):
            return

        try:

            with open(
                self.filename,
                "r",
                encoding="utf-8",
            ) as f:

                self.settings.update(
                    json.load(f)
                )

        except Exception:

            pass

    ##################################################

    def save(self):

        os.makedirs(
            os.path.dirname(self.filename),
            exist_ok=True,
        )

        with open(
            self.filename,
            "w",
            encoding="utf-8",
        ) as f:

            json.dump(
                self.settings,
                f,
                indent=4,
            )

    ##################################################

    def get(
        self,
        key,
        default="",
    ):

        return self.settings.get(
            key,
            default,
        )

    ##################################################

    def set(
        self,
        key,
        value,
    ):

        self.settings[key] = value
        self.save()