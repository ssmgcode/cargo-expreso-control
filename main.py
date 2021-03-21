import click

""" df = pd.read_excel('test.xlsx')
print(df)
print(df['NumeroGuia']) """


@click.command()
def cli():
    print('Hello world from Python CLI!')
