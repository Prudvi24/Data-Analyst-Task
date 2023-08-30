#--------------------- required modules ----------------------
import pandas as pd
import xlsxwriter
import logging
logging.basicConfig(filename="cointab.log",
                    level=logging.INFO,
                    format='%(asctime)s %(name)s %(levelname)s %(message)s')

def read_x_sheets():
    """
    read_x_sheets function loads all the company X related data also called LHS data
    :return:
    """
    try:
        X_order_report = pd.read_excel("/home/prudvi/PycharmProjects/cointab/Assignment details/Company X - Order Report.xlsx")
        X_pincode_zone = pd.read_excel("/home/prudvi/PycharmProjects/cointab/Assignment details/Company X - Pincode Zones.xlsx")
        X_sku_master = pd.read_excel("/home/prudvi/PycharmProjects/cointab/Assignment details/Company X - SKU Master.xlsx")
    except Exception as e:
        logging.exception(e)
    else:
        logging.info("All the company X related data loaded successfully")
        return X_order_report, X_pincode_zone, X_sku_master

def read_courier_sheets():
    """
    read_courier_sheets function loads all the courier company related data also called RHS data
    :return:
    """
    try:
        courier_invoice = pd.read_excel("/home/prudvi/PycharmProjects/cointab/Assignment details/Courier Company - Invoice.xlsx")
        courier_rates = pd.read_excel("/home/prudvi/PycharmProjects/cointab/Assignment details/Courier Company - Rates.xlsx")
    except Exception as e:
        logging.exception(e)
    else:
        logging.info("All the courier company related data loaded successfully")
        return courier_invoice, courier_rates
def merge_sheets(merge_osku, courier_invoice, X_pincode_zone):
    """
    merge_sheets functions takes 3 arguments
    1) merge_osku (After getting total weight associated with each order id)
    2) courier_invoice : Most of the columns are required and important. cols like pincode, billing amount.
    3) pincode zone data : It is used gives us information to which location the order has to be delivered.
                           It specifically extract pincode and zones that are related to the company X
    :param merge_osku:
    :param courier_invoice:
    :param X_pincode_zone:
    :return:
    """
    merged_sheet = pd.merge(courier_invoice,merge_osku, left_on='Order ID', right_on='Order ID')
    merged_sheet.rename(columns={'Charged Weight': 'Total weight as per Courier Company (KG)',
                                 'Billing Amount (Rs.)':'Charges Billed by Courier Company (Rs.)',
                                 'Zone':'Delivery Zone charged by Courier Company'}, inplace=True)
    merged_sheet = pd.merge(merged_sheet,X_pincode_zone, left_on=['Warehouse Pincode','Customer Pincode'], right_on=['Warehouse Pincode','Customer Pincode'])
    merged_sheet.rename(columns={'Zone': 'Delivery Zone as per X'},inplace=True)
    merged_sheet = merged_sheet.drop(['Warehouse Pincode','Customer Pincode'], axis=1)
    logging.info("merged_sheet functions execution completed")
    return merged_sheet

def merge_order_sku(X_order_report, X_sku_master):
    """
    :param X_order_report: company order report
    :param X_sku_master: company SKU data

    The goal of this function, merge_order_sku report is to merge two report to generate
    total weight with respect to the order id.
    which helps in calculation of weight slab related to company X and expected charge.

    :return: merged sheet
    """
    merged_sheet = pd.merge(X_order_report, X_sku_master, left_on='SKU', right_on='SKU')
    logging.info("merge_order_sku function execution completed")
    return merged_sheet

def total_weight_as_per_x(merge_osku):
    """
    :param merge_osku: merged sheet from the function merge_order_sku (merge_osku).

    We calculate total weight associated to each order id in the report sheet.
    Function appends " new columns "Total weight as per X (KG)"

    :return: merge_order_sku sheet having total weight as per X (KG)
    """
    merge_osku['Total weight as per X (KG)'] = merge_osku['Order Qty'] * merge_osku['Weight (g)']
    merge_osku = merge_osku.drop(['SKU', 'Order Qty', 'Item Price(Per Qty.)', 'Weight (g)'], axis=1)
    merge_osku.rename(columns={'ExternOrderNo': 'Order ID'}, inplace=True)
    merge_osku = merge_osku.groupby(['Order ID', 'Payment Mode'])['Total weight as per X (KG)'].sum().reset_index()
    merge_osku['Total weight as per X (KG)'] = merge_osku['Total weight as per X (KG)'].div(1000)
    logging.info("total weight as per x function execution completed")
    return merge_osku

def caculate_expected_charges_by_X(merged_osku_invoice,courier_rates):
    """
    :param merged_osku_invoice: It has all the required information to calculate
                                the weight slab wrt to zone and related to the company X
    :param courier_rates: It is used to give us the information required to calculate the exact expected bill

    As we know bill is calculated with fixed forward charges for the first slab and rest with additional in case of forward charges
    In case of forward and rot. the fixed forward and fixed rto are charged for first slab and rest with additional

    The required logic is implemented below

    :return: dataframe with calculated weight slab as per X expected charge
    """
    # iterate over the merged_osku_involved
    weight_slab_x = []
    charge_by_x = []
    for ind in merged_osku_invoice.index:
        zone = merged_osku_invoice['Delivery Zone as per X'][ind].upper()
        shipment = merged_osku_invoice['Type of Shipment'][ind]
        order_id = merged_osku_invoice['Order ID'][ind]
        ffc = courier_rates.loc[courier_rates['Zone'] == zone, 'Forward Fixed Charge'].values[0]
        fawsc = courier_rates.loc[courier_rates['Zone'] == zone, 'Forward Additional Weight Slab Charge'].values[0]
        rto_fixed = courier_rates.loc[courier_rates['Zone'] == zone, 'RTO Fixed Charge'].values[0]
        rto_add_ws = courier_rates.loc[courier_rates['Zone'] == zone, 'RTO Additional Weight Slab Charge'].values[0]
        weight_slab = courier_rates.loc[courier_rates['Zone']==zone, 'Weight Slabs'].values[0]

        total_weight_x = merged_osku_invoice.loc[merged_osku_invoice['Order ID'] == order_id, 'Total weight as per X (KG)'].values[0]
        ws = total_weight_x / weight_slab
        nws = round(ws,2)
        if (ws - int(ws) == 0):
            nws = int(ws)*weight_slab
            ws = int(ws)
        else:
            nws = (int(ws) + 1)*(weight_slab)
            ws = int(ws)+1
        weight_slab_x.append(nws)

        if(shipment=="Forward charges"):
            charge = 1*ffc + (ws-1)*fawsc
            charge = round(charge,2)
            charge_by_x.append(charge)
        elif(shipment=="Forward and RTO charges"):
            charge = 1*ffc + (ws-1)*fawsc + 1*rto_fixed + (ws-1)*rto_add_ws
            charge = round(charge,2)
            charge_by_x.append(charge)
    merged_osku_invoice['Weight slab as per X (KG)'] = weight_slab_x
    merged_osku_invoice['Expected Charge as per X (Rs.)'] = charge_by_x
    logging.info("calculate expected charges by x is execution completed")
    return merged_osku_invoice

def caculate_weight_slab_courier_company(expected_charges_by_X, courier_rates):
    """
    :param expected_charges_by_X: Expected charge and weight slab as per X datasheet
    :param courier_rates: Sheet that has all the forward and rto charges details

    Since the bill amount given in the sheet
    we are likely to calculate the weight slab as per courier company

    The logic is implemented below
    :return: the dataframe with Weight slab charged by Courier Company (KG) inclued
    """
    weight_slab_courier = []
    for ind in expected_charges_by_X.index:
        zone = expected_charges_by_X['Delivery Zone charged by Courier Company'][ind].upper()
        order_id = expected_charges_by_X['Order ID'][ind]
        weight_slab = courier_rates.loc[courier_rates['Zone'] == zone, 'Weight Slabs'].values[0]

        total_weight_courier = expected_charges_by_X.loc[expected_charges_by_X['Order ID'] == order_id, 'Total weight as per Courier Company (KG)'].values[0]
        ws = total_weight_courier / weight_slab
        if (ws - int(ws) == 0):
            ws = int(ws) * weight_slab
        else:
            ws = (int(ws) + 1) * weight_slab
        weight_slab_courier.append(ws)
    expected_charges_by_X['Weight slab charged by Courier Company (KG)'] = weight_slab_courier
    logging.info("calculate weight slab as per courier company execution completed")
    return expected_charges_by_X

def difference_between_charge_expected(weight_slab_courier_company):
    """
    :param weight_slab_courier_company: Dataframe with all fields required

    1) Difference between the expected charges are calculated.
    A new column is appended to the dataframe "Difference Between Expected Charges and Billed Charges (Rs.)"
    Columns are also ordered according the question / requirement.

    :return: the dataframe "with Difference Between Expected Charges and Billed Charges (Rs.)" column
    """
    weight_slab_courier_company['Difference Between Expected Charges and Billed Charges (Rs.)'] = \
        weight_slab_courier_company['Expected Charge as per X (Rs.)'] - weight_slab_courier_company['Charges Billed by Courier Company (Rs.)']
    weight_slab_courier_company = weight_slab_courier_company.drop(['Type of Shipment', 'Payment Mode'], axis=1)
    # reorder the columns
    weight_slab_courier_company = weight_slab_courier_company[['Order ID', 'AWB Code', 'Total weight as per X (KG)',
                                                               'Weight slab as per X (KG)', 'Total weight as per Courier Company (KG)',
                                                               'Weight slab charged by Courier Company (KG)','Delivery Zone as per X',
                                                               'Delivery Zone charged by Courier Company','Expected Charge as per X (Rs.)',
                                                               'Charges Billed by Courier Company (Rs.)', 'Difference Between Expected Charges and Billed Charges (Rs.)']]
    logging.info("Difference between the expected and billed are calculated successfully")
    return weight_slab_courier_company

def create_summary_sheet(diff_charge_expected):
    """
    :param diff_charge_expected: dataframe having difference of expected chagres with billed charges

    1) overchages, undercharged and correctly charged are calculated
    2) Sum of all charges are also recorded

    :return: summary dataframe with all required manipulation
    """
    overcharged = (diff_charge_expected['Difference Between Expected Charges and Billed Charges (Rs.)'] > 0).sum()
    undercharged = (diff_charge_expected['Difference Between Expected Charges and Billed Charges (Rs.)'] < 0).sum()
    correctly_charged = (diff_charge_expected['Difference Between Expected Charges and Billed Charges (Rs.)'] == 0).sum()

    # Sum positive and negative numbers
    sum_of_overcharged = diff_charge_expected[diff_charge_expected['Difference Between Expected Charges and Billed Charges (Rs.)'] > 0]['Difference Between Expected Charges and Billed Charges (Rs.)'].sum()
    sum_of_negative_charged = diff_charge_expected[diff_charge_expected['Difference Between Expected Charges and Billed Charges (Rs.)'] < 0]['Difference Between Expected Charges and Billed Charges (Rs.)'].sum()
    sum_of_correctly_charged = diff_charge_expected[diff_charge_expected['Difference Between Expected Charges and Billed Charges (Rs.)']==0]['Charges Billed by Courier Company (Rs.)'].sum()


    summary_records = {'Records': ['Total orders where X has been correctly charged','Total Orders where X has been overcharged','Total Orders where X has been undercharged'],
               'Count': [correctly_charged, overcharged, undercharged],
               'Amount (Rs.)': [sum_of_correctly_charged, sum_of_overcharged, sum_of_negative_charged]}
    summary = pd.DataFrame(data=summary_records)
    logging.info("Summary data generated using create_summary sheet function and execution completed")
    return summary

def create_summary_and_calculation(summary_df, diff_charge_expected):
    """
    :param summary_df: summary dataframe to generated summary sheet in the final excel file
    :param diff_charge_expected: calculation dataframe having all calculated fields and required fields

    xlsxwriter engine is used to generate excel sheet having two sheets
    1st sheet is summary sheet
    2nd sheet is calculation sheet

    Final result file is generated in specified path and is named as "result_file.xlsx"
    :return: None
    """
    with pd.ExcelWriter('result_file.xlsx', engine='xlsxwriter') as writer:
        # Save each DataFrame in a separate sheet
        diff_charge_expected['Order ID'] = diff_charge_expected['Order ID'].astype(str)
        diff_charge_expected['AWB Code'] = diff_charge_expected['AWB Code'].astype(str)
        summary_df.to_excel(writer, sheet_name='summary', index=False)
        diff_charge_expected.to_excel(writer, sheet_name='calculations', index=False)
    logging.info("Final result generated in the same path")

logging.info("Logging of cointab begins")
X_order_report, X_pincode_zone, X_sku_master = read_x_sheets()

courier_invoice, courier_rates = read_courier_sheets()

merge_osku = merge_order_sku(X_order_report,X_sku_master)

merge_osku = total_weight_as_per_x(merge_osku)

merged_osku_invoice = merge_sheets(merge_osku,courier_invoice,X_pincode_zone)

expected_charges_by_X = caculate_expected_charges_by_X(merged_osku_invoice,courier_rates)

weight_slab_courier_company = caculate_weight_slab_courier_company(expected_charges_by_X, courier_rates)

diff_charge_expected = difference_between_charge_expected(weight_slab_courier_company);

summary_df = create_summary_sheet(diff_charge_expected)

create_summary_and_calculation(summary_df, diff_charge_expected)
logging.info("Logging of cointab ends")
