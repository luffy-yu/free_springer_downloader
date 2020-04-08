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


def read_and_parse_xlsx(filename) -> dict:
    df = pd.read_excel(filename, sheet_name=sheet_name)
    df = df[[title, doi, eisbn]]
    dic = df.to_dict('records')
    dic = {d[title]: d for d in dic}
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
        filename = name + '.pdf' if not name.endswith('.pdf') else name
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
    dic = {}
    with open(name_eisbn_file, 'r') as f:
        dic = json.loads(f.read())
    found = dic.get(name)
    if not found:
        print('Book "%s" was not found.\n' % name)
        exit(1)
    download(format_download_url(found[doi]), name)


@command.command(help='Download by Xlsx file')
@click.argument('filename', default=free_english_textbooks,
                type=click.Path(file_okay=True, dir_okay=False, exists=True))
def by_xlsx(filename):
    dic = read_and_parse_xlsx(filename)
    names = dic.keys()
    with click.progressbar(names, length=len(names), label='Downloading...') as names:
        for name in names:
            url = format_download_url(dic[name][doi])
            download(url, name)
            time.sleep(1)


if __name__ == '__main__':
    command()
