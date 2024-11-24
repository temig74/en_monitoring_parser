from bs4 import BeautifulSoup
import pandas as pd
import argparse
# import xlsxwriter


def main():
    parser = argparse.ArgumentParser('Process input HTML file and output excel file\nExample: en_mon_parser input.htm output.xlsx\n')
    parser.add_argument('input_filename', help='Path to the input HTML file')
    parser.add_argument('output_filename', help='Path to the output Excel file')
    args = parser.parse_args()
    input_filename = args.input_filename
    output_filename = args.output_filename

    print('Temig monitoring parser 1.0\nParsing html file...')
    with open(input_filename, 'r', encoding="utf-8") as f:
        html = BeautifulSoup(f.read(), 'html.parser')

    rows = []
    for elem in html.find_all('tr'):
        cols = elem.find_all('td')
        if cols[3].text in ('в', 'r') and cols[4].text not in ('Пройден по таймауту', 'Passed by timeout'):
            a = cols[2].find_all('a')
            level = cols[1].text
            player = "'" + a[1].text if a[1].text[0] == "=" else a[1].text
            if len(a[0].text) > 0:
                team = "'"+a[0].text if a[0].text[0] == "=" else a[0].text
            else:
                team = player
            code = "'"+cols[4].text.lower() if cols[4].text[0] == "=" else cols[4].text.lower()
            rows.append([level, team, player, code])

    df = pd.DataFrame(rows, columns=['level', 'team', 'player', 'code'])
    df['level'] = df['level'].astype(int)

    # Количество уникальных кодов, снятых командами на уровнях
    # table1 = df.groupby(['team', 'level'])['code'].nunique().unstack(fill_value=0)
    print('Creating table1')
    table1 = pd.pivot_table(df, values='code', index='team', columns='level', aggfunc='nunique', fill_value=0)
    table1['Total'] = table1.sum(axis=1)
    table1 = table1.sort_values(by=['Total'], ascending=False)

    # Количество уникальных кодов, снятых игроками на уровнях
    print('Creating table2')
    # table2 = df.groupby(['team', 'player', 'level'])['code'].nunique().unstack(fill_value=0)
    table2 = pd.pivot_table(df, values='code', index=['team', 'player'], columns='level', aggfunc='nunique', fill_value=0)
    table2['Total'] = table2.sum(axis=1)

    # считаем сколько и каких команд сняло конкретный код
    print('Creating table3')
    # table3 = df.groupby(['code', 'level', 'team'])['player'].count().unstack(fill_value=0)
    table3 = pd.pivot_table(df, values='player', index=['code', 'level'], columns=['team'], aggfunc='count').fillna(0)
    table3 = (table3 > 0).astype(int).sort_index(level='level')
    table3['Total'] = table3.sum(axis=1)

    print('Writing results to file')
    writer = pd.ExcelWriter(output_filename, engine='xlsxwriter')
    table1.to_excel(writer, sheet_name='Уникальные коды по командам')
    worksheet1 = writer.sheets['Уникальные коды по командам']
    worksheet1.freeze_panes(1, 1)

    table2.to_excel(writer, sheet_name='Уникальные коды по игрокам')
    worksheet2 = writer.sheets['Уникальные коды по игрокам']
    worksheet2.freeze_panes(1, 2)

    table3.to_excel(writer, sheet_name='Коды и команды, их снявшие')
    worksheet3 = writer.sheets['Коды и команды, их снявшие']
    worksheet3.freeze_panes(1, 2)

    writer.close()


if __name__ == '__main__':
    main()
