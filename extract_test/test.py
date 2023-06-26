import camelot
tables = camelot.read_pdf(filepath="pdf\C43011_CH432T_2015-08-21.PDF"
                          ,pages="3")
                        #   ,table_regions=['361,377,518,185'])
camelot.plot(tables[0], kind='line').show()
print(tables)