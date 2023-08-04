import matplotlib.pyplot as plt
import gspread
from oauth2client.service_account import ServiceAccountCredentials


class GoogleSheetGraph:
    def __init__(self, spreadsheet_name, sheet_name):
        self.spreadsheet_name = spreadsheet_name
        self.sheet_name = sheet_name
        self.scope = ['https://spreadsheets.google.com/feeds',
                      'https://www.googleapis.com/auth/spreadsheets',
                      'https://www.googleapis.com/auth/drive.file',
                      'https://www.googleapis.com/auth/drive']
        self.credentials = ServiceAccountCredentials.from_json_keyfile_name('ranking-automation-33939eb3373d.json',
                                                                            self.scope)
        self.client = gspread.authorize(self.credentials)
        self.sheet = self.client.open(self.spreadsheet_name).worksheet(self.sheet_name)

    def fetch_data(self):
        data = self.sheet.get_all_values()
        header = data[0]
        rows = data[1:]
        data_dict = [dict(zip(header, row)) for row in rows]
        return data_dict

    def create_line_graph(self):
        data = self.fetch_data()
        dates = [row['dates'] for row in data]
        rankings = [row['ranking'] for row in data]
        plt.plot(dates, rankings)
        plt.xlabel('Dates')
        plt.ylabel('Rankings')
        plt.title('Rankings over Time')
        plt.show()


if __name__ == "__main__":
    # Graph Creation
    gs_graph = GoogleSheetGraph('Ranking', 'Sheet12')
    gs_graph.create_line_graph()
