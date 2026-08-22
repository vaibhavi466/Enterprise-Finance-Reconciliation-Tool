import os
import pandas as pd

def main():
    os.makedirs('sample_data', exist_ok=True)

    erp_data = pd.DataFrame([
        {'invoice_id': 'INV-1001', 'amount': 1500.00, 'currency': 'USD', 'customer_name': 'Acme Corp', 'date': '2026-08-01'},
        {'invoice_id': 'INV-1002', 'amount': 2500.00, 'currency': 'USD', 'customer_name': 'Beta LLC', 'date': '2026-08-02'},
        {'invoice_id': 'INV-1003', 'amount': 3200.00, 'currency': 'USD', 'customer_name': 'Gamma Inc', 'date': '2026-08-03'},
        {'invoice_id': 'INV-1004', 'amount': 4500.00, 'currency': 'USD', 'customer_name': 'Delta Co', 'date': '2026-08-04'},
        {'invoice_id': 'INV-1005', 'amount': 1200.00, 'currency': 'EUR', 'customer_name': 'Epsilon Ltd', 'date': '2026-08-05'},
        {'invoice_id': 'INV-1007', 'amount': 800.00, 'currency': 'USD', 'customer_name': 'Zeta Enterprises', 'date': '2026-08-07'},
        {'invoice_id': 'INV-1007', 'amount': 800.00, 'currency': 'USD', 'customer_name': 'Zeta Enterprises', 'date': '2026-08-07'},
    ])

    bank_data = pd.DataFrame([
        {'invoice_id': 'INV-1001', 'amount': '$1,500.00', 'currency': 'USD', 'customer_name': 'Acme Corp', 'date': '2026-08-01'},
        {'invoice_id': 'INV-1002', 'amount': '$2,500.004', 'currency': 'USD', 'customer_name': 'Beta LLC', 'date': '2026-08-02'},
        {'invoice_id': 'INV-1003', 'amount': '$3,250.00', 'currency': 'USD', 'customer_name': 'Gamma Inc', 'date': '2026-08-03'},
        {'invoice_id': 'INV-1004', 'amount': '$4,500.00', 'currency': 'EUR', 'customer_name': 'Delta Co', 'date': '2026-08-04'},
        {'invoice_id': 'INV-1006', 'amount': '$5,000.00', 'currency': 'USD', 'customer_name': 'Eta Corp', 'date': '2026-08-06'},
    ])

    erp_data.to_csv('sample_data/erp_data.csv', index=False)
    bank_data.to_excel('sample_data/bank_data.xlsx', index=False, engine='openpyxl')
    print('Sample datasets created successfully in sample_data/')

if __name__ == '__main__':
    main()
