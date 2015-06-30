"""
This script interacts with Defect Dojo products and findings API.
This data is correlated with an external file (CSV) to allow 
mapping of products to their owners.
One Excel file is created with two worksheets:
1. All S0 and S1 findings
2. The total number of S0 and S1 findings per product
"""

import requests
import json
import sqlite3
import glob
import xlsxwriter
from datetime import datetime
import ConfigParser


def get_number_products(url, headers):
    """
    initial API call to get total number of products 
    should using paging someday instead.
    """
    r = requests.get(url,headers=headers) 
    product_data = r.json()
    # grab the total number of products 
    for key, val in product_data.items():
        if key == 'meta':
            product_data_count = val['total_count']
    return product_data_count

def create_system_product_list(url, url_parameters, headers):
    """
    second call to retrieve full list of product names and ids
    returns dictionary of product id and product name
    """
    r = requests.get(url+url_parameters, headers=headers) 
    data = r.json()
    # loop through product list to get names and ids
    for key, val in data.items():
        if key == 'objects':
            products = {}
            for i in range(len(val)):
                products[val[i]['id']] = val[i]['name']
            return products

def create_product_db(file_name):
    """
    create DB for products
    schema includes product name, id and
    dev, QE, SE contacts
    """
    if file_name in glob.glob('*.db'):
        print 'Database file {0} already exists'.format(file_name)
    else:
        print 'Creating database file {0}'.format(file_name)
        conn = sqlite3.connect(file_name)
        curs = conn.cursor()
        curs.execute('''CREATE TABLE products
                     (id integer primary key, dojo_id, dojo_name, dev_mgr, qe_mgr, se_team)''') 
        conn.commit()
        conn.close()
    return True

def populate_db_products(product_dictionary, file_name):
    """
    insert product data into database from dojo product url 
    """
    conn = sqlite3.connect(file_name)
    curs = conn.cursor()
    existing_db = read_db_products(file_name)
    # if nothing in the database, just insert all products
    if len(existing_db) == 0:
        print 'Populating database for first time'
        for key, value in product_dictionary.items():
            curs.execute("INSERT INTO products(dojo_id, dojo_name) VALUES (?, ?)", (key, value))
    # the database has entries, so compare latest list to existing entries
    # only insert new products
    else:
        for db_row in existing_db:
            for key, value in product_dictionary.items():
                if key in db_row:
                    break
            else:
                print 'Inserting new product: {0}'.format(value)
                curs.execute("INSERT INTO products(dojo_id, dojo_name) VALUES (?, ?)", (key, value))
    conn.commit()
    #for row in conn.execute("SELECT * FROM products"):
    #    print row
    conn.close()
    return True

def populate_db_owners(database_file, owners_file):
    """
    add product owner data to database from mapping file
    """
    cnt = 0
    f = open(owners_file)
    conn = sqlite3.connect(database_file)
    curs = conn.cursor()
    for row in curs.execute("SELECT * FROM products"):
        for line in f:
            if str(row[2]) in line.split(',')[0]:
                curs2 = conn.cursor()
                cnt += 1
                #print '{0} was found in {1}! Inserting into database'.format(row[2], line)
                product_id = row[1]
                dev = line.split(',')[1].replace('\n', '')
                qe = line.split(',')[2].replace('\n', '')
                se = line.split(',')[3].replace('\n', '')
                #print ' Inserting data {0} : {1} : {2} : {3} : {4}'.format(row[2], product_id, dev, qe, se)
                curs2.execute("UPDATE products SET dev_mgr = ?, qe_mgr = ?, se_team = ? \
                                WHERE dojo_id = ?", (dev, qe, se, product_id))
        f.seek(0)
    conn.commit()
    conn.close()
    print 'Read {0} owners from {1}. Database updated with any changes.'.format(cnt, owners_file)
    return True   

def read_db_products(file_name):
    """
    read from product database
    returns list of tuples. each tuple is a db row
    """
    db_product_list = []
    conn = sqlite3.connect(file_name)
    curs = conn.cursor()
    for row in conn.execute("SELECT * FROM products"):
        db_product_list.append(row)
    return db_product_list

def compare_system_db_products(system_list, db_list):
    """
    TO DO: may not need this due to populate_db_products
    compare the system products to
    database products to find new products
    """
    print "system {0} : db {1}".format(len(system_list), len(db_list))
    if len(system_list) == len(db_list):
        print 'No product changes detected'
        # TO DO:
        # Actually compare list by product ID and
        # check for name changes or 
        # additions + removals
    else:
        print 'New products found!'
        # TO DO: 
        # find the new products and 
        # load them into database

def get_number_findings(url, headers):
    """
    API call to get number of findings
    should use paging instead.
    """
    r = requests.get(url,headers=headers) 
    finding_data = r.json()
    # grab the total number of findings 
    for key, val in finding_data.items():
        if key == 'meta':
            finding_data_count = val['total_count']
    return finding_data_count

def create_system_finding_list(url, url_parameters, headers):
    """
    call to retrieve full list of findings
    returns list of dictionaries (json)
    """
    r = requests.get(url+url_parameters, headers=headers) 
    data = r.json()
    # loop through finding list to get info
    for key, val in data.items():
        if key == 'objects':
            findings = []
            for i in range(len(val)):
                findings.append({'title':val[i]['title'],
                                'severity': val[i]['numerical_severity'],
                                'product' : val[i]['product'],
                                'date' : val[i]['date']})
            return findings

def group_findings(product_list, finding_list):
    """
    group findings by product
    return list of finding data as strings
    """
    today = datetime.today()
    report = []
    for k in range(len(product_list)):
#       print product_list[k][1]
        for j in range(len(finding_list)):
            #for keyJ, valJ in finding_list[j].items():
            if str(product_list[k][1]) in finding_list[j]['product']:
                    report.append('{0},{1},{2},{3},{4}'.format(product_list[k][2], finding_list[j]['title'], \
                                    finding_list[j]['date'], finding_list[j]['severity'], \
                                    (today - datetime.strptime(finding_list[j]['date'], '%Y-%m-%d')).days ))
    return report

def count_findings(product_list, finding_list):
    """
    count S0 & S1 findings by product
    return list of dictionaries 
    """
    report = []
    for k in range(len(product_list)):
        report.append({product_list[k][1] : {'S0' : 0, 'S1' : 0, 'name' : product_list[k][2], \
                                                'dev' : product_list[k][3], 'qe' : product_list[k][4], \
                                                'se' : product_list[k][5]}})
        for j in range(len(finding_list)):
            #for keyJ, valJ in finding_list[j].items():
            if str(product_list[k][1]) in finding_list[j]['product']:
                if finding_list[j]['severity'] == 'S0':
                    report[k][product_list[k][1]]['S0'] += 1
                else:
                    report[k][product_list[k][1]]['S1'] += 1
    return report

def create_metrics_report(findings):
    """
    returns list of products with totalled S0 & S1 - as strings 
    """
    report = []
    for x in range(len(findings)):
        for key, val in findings[x].items():
            line = '{0},{1},{2},{3},{4},{5}'.format(val['name'],val['dev'],val['qe'],val['se'],val['S0'],val['S1'])
            report.append(line)
    #f.close
    return report

def create_report():
    config = ConfigParser.ConfigParser()
    config.read("./config.ini")
    headers = {'content-type': config.get("header","contenttype"),
            'Authorization': config.get("header","Authorization")} 
    # get total number of products from dojo
    product_url = config.get("url", "product")
    num_products = get_number_products(product_url, headers)
    # set limit parameter to the total number of products 
    limit_parameter = '?limit={0}'.format(num_products)
    # get list of all products from dojo
    system_product_list = create_system_product_list(product_url, limit_parameter, headers)
    # create the products and owners database
    database_file_name = config.get("file", "database")
    database_file = create_product_db(database_file_name)
    # populate the database with product info
    full_database = populate_db_products(system_product_list, database_file_name)
    # this returns the products database in a list of tuples.
    db_product_list = read_db_products(database_file_name)
    # add owners to the database from a mapping CSV file
    mapping_file_name = config.get("file", "mapping")
    pop = populate_db_owners(database_file_name, mapping_file_name)
    # get total number of findings from dojo
    finding_url = config.get("url", "finding")
    findings_total = get_number_findings(finding_url, headers)
    # create parameter to allow all S0 & S1 findings to be returned in one call
    finding_parameter = '?active=true&verified=true&severity__in=Critical,High&limit={0}'.format(findings_total)
    # get list of dicts - all S0 & S1 findings in dojo.
    all_findings = create_system_finding_list(finding_url, finding_parameter, headers)
    all_findings.sort()
    today = str(datetime.now().date())
    report_name = config.get("file", "spreadsheet")
    full_report_name = report_name + '_' + today
    workbook = xlsxwriter.Workbook(full_report_name + '.xlsx')
    # create report header formatting
    # white font; red background
    ws_format = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#800000'})
    # create date column format for report
    date_format = workbook.add_format({'num_format': 'mmm dd, yyyy'})
    # start first worksheet for findings report
    finding_sheet_name = config.get("file", "finding")
    f_sheet = workbook.add_worksheet(finding_sheet_name)
    # set column widths for report
    # column start, column stop, width
    f_sheet.set_column(0, 0, 22)
    f_sheet.set_column(1, 1, 80)
    f_sheet.set_column(2, 2, 13)
    f_sheet.set_column(3, 4, 8)
    # create column headers
    f_sheet.write('A1', 'Product', ws_format)
    f_sheet.write('B1', 'Title of Finding', ws_format)
    f_sheet.write('C1', 'Identified Date', ws_format)
    f_sheet.write('D1', 'Severity', ws_format)
    f_sheet.write('E1', 'Age (days)', ws_format)
    # increment data fields through worksheet
    # start row at 1 to leave the headers we created
    row = 1
    col = 0
    findings_finding_list = group_findings(db_product_list, all_findings)
    for index in range(len(findings_finding_list)):
        f_sheet.write(row, col, findings_finding_list[index].split(',')[0])
        f_sheet.write(row, col + 1, findings_finding_list[index].split(',')[1])
        age = datetime.strptime(findings_finding_list[index].split(',')[2], '%Y-%m-%d')
        f_sheet.write_datetime(row, col + 2, age, date_format)
        f_sheet.write(row, col + 3, findings_finding_list[index].split(',')[3])
        f_sheet.write_number(row, col + 4, int(findings_finding_list[index].split(',')[4]))
        row += 1
    # start second worksheet for metrics report
    metric_sheet_name = config.get("file", "metrics")
    m_sheet = workbook.add_worksheet(metric_sheet_name)
    # set column widths for report
    # column start, column stop, width
    m_sheet.set_column(0, 0, 26)
    m_sheet.set_column(1, 2, 15)
    m_sheet.set_column(3, 3, 8)
    m_sheet.set_column(4, 5, 16)
    # create column headers
    m_sheet.write('A1', 'Product', ws_format)
    m_sheet.write('B1', 'Dev Mgr', ws_format)
    m_sheet.write('C1', 'QE Mgr', ws_format)
    m_sheet.write('D1', 'SE Mgr', ws_format)
    m_sheet.write('E1', 'Number of Active S0', ws_format)
    m_sheet.write('F1', 'Number of Active S1', ws_format)
    # increment data fields through worksheet
    # start row at 1 to leave the headers we created
    row = 1
    col = 0
    m_findings = count_findings(db_product_list, all_findings)
    metrics_finding_list = create_metrics_report(m_findings)
    metrics_finding_list.sort()
    for index in range(len(metrics_finding_list)):
        m_sheet.write(row, col, metrics_finding_list[index].split(',')[0])
        m_sheet.write(row, col + 1, metrics_finding_list[index].split(',')[1])
        m_sheet.write(row, col + 2, metrics_finding_list[index].split(',')[2])
        m_sheet.write(row, col + 3, metrics_finding_list[index].split(',')[3])
        m_sheet.write_number(row, col + 4, int(metrics_finding_list[index].split(',')[4]))
        m_sheet.write_number(row, col + 5, int(metrics_finding_list[index].split(',')[5]))
        row += 1
    workbook.close()
    workbook_name = full_report_name + '.xlsx'
    return workbook_name

# Create finished XLS report
metrics_report = create_report()
print 'Report {0} created.'.format(metrics_report)
