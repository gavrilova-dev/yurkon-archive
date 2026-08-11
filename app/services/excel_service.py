from openpyxl import load_workbook

from app.models.registry_entry import RegistryEntry


class ExcelService:

    def __init__(self, path):

        self.path = path
        self.records = []
        self.workbook = None
        self.sheet = None


    def load(self):

        self.workbook = load_workbook(
            self.path
        )

        self.sheet = self.workbook.active

        self.records = []

        headers = {}

        # Заголовки во 2 строке
        for cell in self.sheet[2]:

            if cell.value:

                name = str(cell.value).strip()

                headers[name] = cell.column


        protocol_col = headers.get(
            "Номер протокола"
        )

        index_col = headers.get(
            "Индекс дела"
        )

        organization_col = headers.get(
            "Наименование организации"
        )


        if not protocol_col:

            raise Exception(
                "Не найден столбец 'Номер протокола'"
            )


        if not index_col:

            raise Exception(
                "Не найден столбец 'Индекс дела'"
            )


        if not organization_col:

            raise Exception(
                "Не найден столбец 'Наименование организации'"
            )


        for row in range(
            3,
            self.sheet.max_row + 1
        ):

            protocol = self.sheet.cell(
                row=row,
                column=protocol_col
            ).value


            index = self.sheet.cell(
                row=row,
                column=index_col
            ).value


            organization = self.sheet.cell(
                row=row,
                column=organization_col
            ).value


            if protocol and index:

                self.records.append(
                    RegistryEntry(
                        protocol=str(protocol),
                        index=str(index),
                        organization=str(organization),
                        excel_row=row,
                        scanned=False
                    )
                )


    def find_by_protocol(self, protocol):

        protocol = str(
            protocol
        ).strip().lower()


        if not protocol:

            return None


        matches = []


        for record in self.records:

            excel_protocol = str(
                record.protocol
            ).strip().lower()


            # Если пользователь ввёл часть номера:
            # 9750 → найдёт 9750/2025,
            # 9750/1 → найдёт 9750/1/2025.
            if excel_protocol.startswith(protocol):

                matches.append(record)


        if not matches:

            return None


        # Одно совпадение
        if len(matches) == 1:

            return matches[0]


        # Несколько совпадений
        return matches