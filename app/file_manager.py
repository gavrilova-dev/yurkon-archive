from pathlib import Path


class FileManager:

    def __init__(
        self,
        incoming_folder,
        ready_folder
    ):

        self.incoming_folder = Path(incoming_folder)
        self.ready_folder = Path(ready_folder)


        self.incoming_folder.mkdir(
            exist_ok=True
        )

        self.ready_folder.mkdir(
            exist_ok=True
        )



    def get_new_files(self):

        return sorted(
            self.incoming_folder.glob("*.pdf")
        )



    def move_pdf(
        self,
        pdf_path,
        index
    ):

        filename = f"{index.replace('/', '-')}.pdf"


        destination = (
            self.ready_folder /
            filename
        )


        if destination.exists():

            raise Exception(
                f"Файл уже существует:\n{destination.name}"
            )


        pdf_path.rename(
            destination
        )


        return destination