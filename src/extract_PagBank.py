from xml.etree import ElementTree as ET
import pandas as pd

def extract_ofx_data(file_path):
    # Leitura do arquivo OFX como texto com codificação UTF-8
    with open(file_path, 'r', encoding='utf-8') as file:
        ofx_content = file.read()

    # Tratamento do conteúdo OFX para XML parsing
    ofx_content = ofx_content.replace('\n', '').replace('\r', '')
    ofx_content = ofx_content[ofx_content.find('<OFX>'):]

    # Parse do conteúdo OFX
    root = ET.fromstring(ofx_content)

    # Extração de informações da conta
    bank_id = root.findtext('.//BANKID')
    account_id = root.findtext('.//ACCTID')
    account_type = root.findtext('.//ACCTTYPE')

    # Função para limpar e converter o saldo
    def clean_balance(balance_str):
        if balance_str:
            balance_str = balance_str.replace('R$', '').replace('.', '').replace(',', '.').replace('\xa0', '').strip()
            return float(balance_str)
        return None

    # Extração de saldo inicial e final
    ledger_balance = clean_balance(root.findtext('.//LEDGERBAL/BALAMT'))
    available_balance = clean_balance(root.findtext('.//AVAILBAL/BALAMT'))

    # Extração de datas de início e fim
    start_date = root.findtext('.//BANKTRANLIST/DTSTART')
    end_date = root.findtext('.//BANKTRANLIST/DTEND')

    # Função para converter datas no formato OFX
    def convert_ofx_date(ofx_date):
        cleaned_date = ofx_date[:14]  # Pega apenas "20240819172247"
        return pd.to_datetime(cleaned_date, format='%Y%m%d%H%M%S')

    # Extração das transações
    transactions = []

    for transaction in root.findall('.//STMTTRN'):
        trn_type = transaction.findtext('TRNTYPE')
        date = transaction.findtext('DTPOSTED')
        amount = transaction.findtext('TRNAMT')
        fitid = transaction.findtext('FITID')
        payee = transaction.findtext('NAME')
        memo = transaction.findtext('MEMO')

        transactions.append({
            'Tipo de Transação': trn_type,
            'Data': convert_ofx_date(date),
            'Valor': float(amount),
            'ID da Transação': fitid,
            'Beneficiário': payee,
            'Descrição': memo
        })

    # Convertendo para DataFrame para melhor visualização
    transactions_df = pd.DataFrame(transactions)

    # Resumo da conta separado por informações
    agency = bank_id
    account_number = account_id
    account_type = account_type
    initial_balance = ledger_balance
    available_balance = available_balance
    start_date = convert_ofx_date(start_date) if start_date else None
    end_date = convert_ofx_date(end_date) if end_date else None

    return transactions_df, agency, account_number, account_type, initial_balance, available_balance, start_date, end_date
