import pandas as pd
from sqlalchemy import create_engine
from gooey import Gooey, GooeyParser
from fuzzywuzzy import fuzz

titles = ['นาย', 'นางสาว', 'นาง', 'นส', 'ดร', 'ทนพ', 'ทนพญ']

def check_member_license(excel_file, db, lic_col, fname_col, lname_col):
    engine = create_engine('sqlite:///{}'.format(db))
    data = pd.read_excel(excel_file)
    def clean_name(row):
        if fname_col != lname_col:
            name = '{} {}'.format(row[fname_col - 1], row[lname_col - 1])
        else:
            name = row[fname_col - 1]
        if pd.isnull(name):
            return ''

        name = name.strip().replace('.', '')
        for t in titles:
            name = name.replace(t, '')
        return name

    def check_license_id(row):
        license_id = row[lic_col - 1]
        if not pd.isnull(license_id):
            print(f'Checking {license_id}')
            name = row['cleaned_name']
            rec = df.query(f'mem_id == {license_id}')
            if not rec.empty:
                fullname = '{} {}'.format(rec.iloc[0]['fname'],
                                          rec.iloc[0]['lname'])
                e_fullname = '{} {}'.format(rec.iloc[0]['e_fname'],
                                          rec.iloc[0]['e_lname'])
                row['db_th_name'] = fullname
                row['db_en_name'] = e_fullname
                th_fullname_ratio = fuzz.ratio(fullname, name)
                en_fullname_ratio = fuzz.ratio(e_fullname.upper(), name.upper())
                row['ratio'] = th_fullname_ratio if th_fullname_ratio > en_fullname_ratio \
                    else en_fullname_ratio
                return row
            else:
                row['db_th_name'] = 'Not found'
                row['db_en_name'] = 'Not found'
                row['ratio'] = 0
                return row
        else:
            row['db_th_name'] = 'No license'
            row['db_en_name'] = 'No license'
            row['ratio'] = None
            return row

    df = pd.read_sql_table('members', con=engine)
    data['cleaned_name'] = data.apply(clean_name, axis=1)
    data = data.apply(check_license_id, axis=1)
    data.to_excel('output.xlsx', index=False)
    print('Total processed records is {}.'.format(len(data)))


@Gooey
def main():
    parser = GooeyParser(description='CMTE License No. Checker')
    parser.add_argument('filename', widget='FileChooser')
    parser.add_argument('db', widget='FileChooser')
    parser.add_argument('license_col', type=int, help='Column number of the license number')
    parser.add_argument('fname_col', type=int, help='Column number of the first name')
    parser.add_argument('lname_col', type=int, help='Column number of the last name')
    args = parser.parse_args()
    check_member_license(args.filename,
                         args.db,
                         args.license_col,
                         args.fname_col,
                         args.lname_col)

if __name__ == '__main__':
    main()