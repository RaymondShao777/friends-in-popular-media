import csv, xlsxwriter

def write_xlsx(csv_file_name):
    workbook = xlsxwriter.Workbook("wsj.xlsx")
    HEADER = ['Headline', 'Link', 'Date', 'Summary']
    BOLD = workbook.add_format({'bold':True})
    WRAP = workbook.add_format()
    WRAP.set_shrink()


    # set up excel sheet
    cur_row = {}
    worksheets = {}
    year = 2015
    for i in range(10):
        worksheets[str(year)] = workbook.add_worksheet(str(year))
        worksheets[str(year)].write_row(0, 0, HEADER, BOLD)
        cur_row[str(year)] = 1
        year += 1


    with open(csv_file_name) as csv_file:
        reader = csv.DictReader(csv_file)
        for row in reader:
             year = row['PubDate'].strip()[:4]
             if not year:
                 continue
             row_to_write = (row['Title'], row['DocumentUrl'], row['PubDate'], row['Abstract'])

             worksheets[year].write_row(cur_row[year], 0, row_to_write, WRAP)
             cur_row[year] += 1

    workbook.close()


# HEADLINE, LINK, DATE, SUMMARY
write_xlsx("wsj_export.csv")
