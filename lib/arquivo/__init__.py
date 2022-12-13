import pandas as pd


def read_excel(arquivo):
    """
    Abri arquivo xlsx utilizando pandas
    :param arquivo: xlsx
    :return: list (Nome, PIS e Quantidade)
    """
    df = pd.read_excel(arquivo)
    df['PIS'] = df['PIS'].apply(str)
    df['Nome'] = df['Nome'].apply(str.upper)
    return list(df['Nome']), list(df['PIS']), len(list(df['PIS']))


if __name__ == '__main__':
    pass
