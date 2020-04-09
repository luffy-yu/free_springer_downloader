import json
import os
import time

import click
import pandas as pd
import wget

free_english_textbooks = 'Free+English+textbooks.xlsx'
url_template = 'https://link.springer.com/content/pdf/{DOI}.pdf'

name_eisbn_file = 'name_eisbn.json'

title = 'Book Title'
doi = 'DOI URL'
eisbn = 'Electronic ISBN'
sheet_name = 'eBook list'

"""
Duplicate Book Names:

('Ceramic Materials', 2)
('Fundamentals of Biomechanics', 2)
('Transmission Electron Microscopy', 2)
('Introduction to Partial Differential Equations', 2)
('Introduction to Logic Circuits & Logic Design with Verilog', 2)
('Additive Manufacturing Technologies', 2)
('Computational Physics', 2)
('Fundamentals of Business Process Management', 2)
('Strategic International Management', 2)
('Robotics, Vision and Control', 2)
('Probability Theory', 2)
('Quantum Mechanics', 2)
('Robotics', 2)
('Pharmaceutical Biotechnology', 2)
('Introduction to Logic Circuits & Logic Design with VHDL ', 2)
('Advanced Organic Chemistry', 2)
"""

duplicate_names = {'Ceramic Materials', 'Fundamentals of Biomechanics', 'Transmission Electron Microscopy',
                   'Introduction to Partial Differential Equations',
                   'Introduction to Logic Circuits & Logic Design with Verilog', 'Additive Manufacturing Technologies',
                   'Computational Physics', 'Fundamentals of Business Process Management',
                   'Strategic International Management', 'Robotics, Vision and Control', 'Probability Theory',
                   'Quantum Mechanics', 'Robotics', 'Pharmaceutical Biotechnology',
                   'Introduction to Logic Circuits & Logic Design with VHDL ', 'Advanced Organic Chemistry'}


def read_and_parse_xlsx(filename) -> list:
    df = pd.read_excel(filename, sheet_name=sheet_name)
    df = df[[title, doi, eisbn]]
    dic = df.to_dict('records')
    # avoid duplicate names
    # dic = {d[title]: d for d in dic}
    return dic


def prepare():
    if not os.path.exists(name_eisbn_file):
        if os.path.exists(free_english_textbooks):
            dic = read_and_parse_xlsx(free_english_textbooks)
            with open(name_eisbn_file, 'w', encoding='utf-8') as f:
                f.write(json.dumps(dic, indent=2, ensure_ascii=False))
        else:
            print('File "%s" was not found.\n' % free_english_textbooks)
            exit(1)


def format_download_url(doi):
    return url_template.format(DOI=doi.replace('http://doi.org/', ''))


def download(url, name):
    try:
        # replace '[:/]' by '_'
        name = name.replace(':', '_').replace('/', '_')
        filename = name + '.pdf' if not name.endswith('.pdf') else name
        if os.path.exists(filename) and name not in duplicate_names:
            return
        wget.download(url, filename)
    except:
        with open('log.log', 'a') as f:
            f.write('Failed to download "%s" .\n' % name)


@click.group()
def command():
    pass


@command.command(help='Download by Name')
@click.argument('name')
def by_name(name):
    prepare()
    dic = []
    with open(name_eisbn_file, 'r') as f:
        dic = json.loads(f.read())
    found = list(filter(lambda x: x[title] == name, dic))
    if not found:
        print('Book "%s" was not found.\n' % name)
        exit(1)
    [download(format_download_url(one[doi]), name) for one in found]


@command.command(help='Download by Xlsx file')
@click.argument('filename', default=free_english_textbooks,
                type=click.Path(file_okay=True, dir_okay=False, exists=True))
def by_xlsx(filename):
    dic = read_and_parse_xlsx(filename)
    with click.progressbar(dic, length=len(dic), label='Downloading...') as books:
        for book in books:
            url = format_download_url(book[doi])
            download(url, book[title])
            time.sleep(1)


if __name__ == '__main__':
    command()
