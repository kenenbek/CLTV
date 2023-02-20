from ololo import Metrics, SpreadSheet

if __name__ == '__main__':
    SAMPLE_SPREADSHEET_ID = '1NnE5eDIe3w3HWa9iyl1t-xOUI2MuuiDFcr_rcL1aeH8'
    SAMPLE_RANGE_NAME = '01/2022!B11:B'  # example of format

    MONTH_S = ["01/2022", "02/2022", "03/2022", "04/2022",
               "05/2022", "06/2022", "07/2022", "08/2022",
               "09/2022", "10/2022", "11/2022", "12/2022",
               "01/2023"]

    MONEY_COLUMN = ["I", "I", "I", "G",
                    "G", "G", "G", "G",
                    "G", "G", "G", "G",
                    "G"]

    spreadsheet = SpreadSheet()
    metrics = Metrics()

    company_office_amount_dict = spreadsheet.make_request_two_column(SAMPLE_SPREADSHEET_ID,
                                                                     'Price policy!B4:B', 'Price policy!G4:G')
    company_office_amount_dict = metrics.clean_company_people_amount(company_office_amount_dict)

    for money_column, month in zip(MONEY_COLUMN, MONTH_S):
        client_money_one_month_dict = None
        try:
            client_money_one_month_dict = spreadsheet.make_request_two_column(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                                                              range1="{}!B11:B".format(month),
                                                                              range2="{}!{}11:{}".format(month,
                                                                                                         money_column,
                                                                                                         money_column))
        except Exception as e:
            print("Bad" + "    " + month)
            print(str(e))

        client_money_one_month_dict = metrics.clean_client_money_one_month(client_money_one_month_dict)
        metrics.update_client_duration(client_money_one_month_dict)
        metrics.update_income_from_clients(client_money_one_month_dict)

        # print(values)

    metrics.sort_duration_dict()
    metrics.sort_clients_duration_money_dict()

    import numpy as np
    for k, v in metrics.client_incomes.items():
        print(k, metrics.company_people_amount_dict.get(k, 1), np.sum(v) / metrics.company_people_amount_dict.get(k, 1), sep=", ")