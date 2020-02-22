import WaApi
import urllib.parse
import csv
import argparse
import os
import json
from flask import Flask, request
from intuitlib.client import AuthClient
from intuitlib.enums import Scopes
from quickbooks import QuickBooks
from quickbooks.objects.customer import Customer, Address, PhoneNumber, EmailAddress
from quickbooks.objects.payment import Payment, PaymentLine
from quickbooks.objects.salesreceipt import SalesReceipt
from quickbooks.objects.refundreceipt import RefundReceipt
from quickbooks.objects.account import Account
from quickbooks.objects.item import Item
from quickbooks.objects.paymentmethod import PaymentMethod
from quickbooks.objects.base import Ref
from quickbooks.objects.detailline import SalesItemLine, SalesItemLineDetail
from quickbooks.objects.trackingclass import Class

from  QBO_Credentials import *

app = Flask(__name__)

#CLIENT_ID='AB5wgTQYXXrenxcNxqpBlucUjiQQGimvWCLNQQXAtzKFlHbQNb'
#CLIENT_SECRET='Do9JDFZvMf2RdA75Tq3asQSi0vSeGXFESaau6Ikg'
#ENVIRONMENT='sandbox'
#REDIRECT_URI='http://localhost:5000/callback'

#CLIENT_ID='AB2VGBhHkMjjmtddNqzuVqW7im8fcwsvVes7r9MNkZdAcKGlkr'
#CLIENT_SECRET='IRz8yE9WKEeRPbf3POXUC4cJqCHNHWanURfNMKTQ'
#ENVIRONMENT='production'
#REDIRECT_URI='https://local.roseundy.net:5000/callback'

STATE_TOKEN = "deadbeef"

QB_WA_PRODUCT_CLASSES = "Hub Classes"
QB_WA_PRODUCT_SUMMER_CAMPS = "Summer Camp"
QB_WA_PRODUCT_MEMBERSHIPS = "Membership Dues"
QB_WA_PRODUCT_GIFTCERTIFICATE = "Gift Certificate"
QB_WA_PRODUCT_REFUNDS = "Refund"
QB_WA_PRODUCT_EXCESS = "Overpayment"
QB_WA_PRODUCT_UNDEFINED = "General"

QB_WA_CLASS_CLASSES = "Classes"
QB_WA_CLASS_SUMMER_CAMPS = "Summer Camp"
QB_WA_CLASS_MEMBERSHIPS = "Maker Space"
QB_WA_CLASS_GIFTCERTIFICATE = "Maker Space"
QB_WA_CLASS_REFUNDS = "Maker Space"
QB_WA_CLASS_EXCESS = "Maker Space"
QB_WA_CLASS_UNDEFINED = "Maker Space"

# I'm going to have to fix this
QB_WA_CUSTOMER_CLASSES = "Attendee"
QB_WA_CUSTOMER_REFUNDS = "Wild Apricot"
QB_WA_CUSTOMER_EXCESS = "Wild Apricot"

tender2dest_payment = {
                'Cash': 'Undeposited Funds',
                'Check': 'Undeposited Funds',
                'Credit Card': 'AffiniPay',
                'Payline': 'Payline',
                'PayPal': 'PayPal',
                'Gift Certificate': 'Gift Certificates Outstanding',
                'Wire transfer': 'N/A',
                'Special discount': 'N/A',
                'In-kind payment': 'N/A',
                'PayPal Payments Standard': 'N/A',
                'PayPal Subscription Payment': 'N/A',
                'PayPal Credit Card': 'N/A',
                'PayPal Credit Card Recurring Payment': 'N/A',
                'PayPal Express Checkout': 'N/A',
                'PayPal Recurring Payment': 'N/A',
                'PayPal Payments Advanced': 'N/A',
                'PayPal Payflow Pro Credit Card': 'N/A',
                'PayPal Payflow Pro Recurring Payment': 'N/A',
                'Authorize.NET Credit Card': 'N/A',
                'Authorize.NET Recurring Payment': 'N/A',
                'Google Account': 'N/A',
                '2Checkout': 'N/A',
                '2Checkout Recurring Payment': 'N/A',
                'BluePay Credit Card': 'N/A',
                'BluePay Recurring Payment': 'N/A',
                'CRE Secure Credit Card': 'N/A',
                'CRE Secure Recurring Payment': 'N/A',
                'Global Payments Credit Card': 'N/A',
                'IATS Credit Card': 'N/A',
                'IATS Recurring Payment': 'N/A',
                'Moneris Credit Card': 'N/A',
                'Moneris Recurring Payment': 'N/A',
                'Skrill': 'N/A',
                'Skrill Recurring Payment': 'N/A',
                'Stripe': 'N/A',
                'Stripe Recurring Payment': 'N/A',
                'Wild Apricot Payment': 'AffiniPay',
                'Wild Apricot Recurring Payment': 'AffiniPay',
    }

tender2dest_refund = {
                'Cash': 'Undeposited Funds',
                'Check': 'Checking - General (0644)',
                'Credit Card': 'AffiniPay',
                'Payline': 'Payline',
                'PayPal': 'PayPal',
                'Gift Certificate': 'Gift Certificates Outstanding',
                'Wire transfer': 'N/A',
                'Special discount': 'N/A',
                'In-kind payment': 'N/A',
                'PayPal Payments Standard': 'N/A',
                'PayPal Subscription Payment': 'N/A',
                'PayPal Credit Card': 'N/A',
                'PayPal Credit Card Recurring Payment': 'N/A',
                'PayPal Express Checkout': 'N/A',
                'PayPal Recurring Payment': 'N/A',
                'PayPal Payments Advanced': 'N/A',
                'PayPal Payflow Pro Credit Card': 'N/A',
                'PayPal Payflow Pro Recurring Payment': 'N/A',
                'Authorize.NET Credit Card': 'N/A',
                'Authorize.NET Recurring Payment': 'N/A',
                'Google Account': 'N/A',
                '2Checkout': 'N/A',
                '2Checkout Recurring Payment': 'N/A',
                'BluePay Credit Card': 'N/A',
                'BluePay Recurring Payment': 'N/A',
                'CRE Secure Credit Card': 'N/A',
                'CRE Secure Recurring Payment': 'N/A',
                'Global Payments Credit Card': 'N/A',
                'IATS Credit Card': 'N/A',
                'IATS Recurring Payment': 'N/A',
                'Moneris Credit Card': 'N/A',
                'Moneris Recurring Payment': 'N/A',
                'Skrill': 'N/A',
                'Skrill Recurring Payment': 'N/A',
                'Stripe': 'N/A',
                'Stripe Recurring Payment': 'N/A',
                'Wild Apricot Payment': 'N/A',
                'Wild Apricot Recurring Payment': 'N/A',
    }

dues2customer = {
    22.5: 'Student/Senior Subscription',
    45: 'Individual Subscription',
}

apiKey_fname = 'client_secret'
bin_dir = os.path.dirname(__file__)
app_dir = os.path.dirname(bin_dir)
apiKey_fpath = os.path.join(
    #app_dir,
    bin_dir,
    'etc',
    apiKey_fname)

def get_apiKey(kpath):
    """Reads Wild Apricot API key from a file

    Returns: api key string
    """
    with open(kpath, 'r') as f:
        apiKey = f.readline().strip()
    return (apiKey)

def get_contacts(api, debug, contactsUrl):
    """Make an API call to Wild Apricot to retrieve
    contact info for all contacts

    Returns: list of contacts
    """
    #params = {'$filter': 'member eq true AND Status eq Active',
    #          '$async': 'false'}
    params = {'$async': 'false'}
    request_url = contactsUrl + '?' + urllib.parse.urlencode(params)
    if debug: print('Making api call to get all contacts')
    return api.execute_request(request_url).Contacts

# create hash of id -> contact index
def make_contact_hash (contacts):
    cnt = 0
    chash = {}
    for c in contacts:
        chash[c.Id] = cnt
        cnt += 1
    return chash


def get_payments (api, debug, paymentsUrl, start, stop):
    params = {'StartDate': start,
              'EndDate'  : stop}
    request_url = paymentsUrl + '?' + urllib.parse.urlencode(params)
    if debug: print('Making api call to get payments')
    return api.execute_request(request_url)    

def get_warefunds (api, debug, refundsUrl, start, stop):
    params = {'StartDate': start,
              'EndDate'  : stop}
    request_url = refundsUrl + '?' + urllib.parse.urlencode(params)
    if debug: print('Making api call to get refunds')
    return api.execute_request(request_url)    

# dump out sales array into a string
def dump_wasales (sales):
    outstr = ""
    for s in sales:
        outstr += 'Payment Id: ' + str(s['payment_id']) + '\n'
        outstr += 'Date: ' + s['date'] + '\n'
        outstr += 'Type: ' + s['type'] + '\n'
        outstr += 'Contact: ' + s['contact_name'] + ' ('+str(s['contact_id'])+')' + '\n'
        outstr += 'Level: ' + s['level'] + '\n'
        outstr += 'Tender: ' + s['tender'] + '\n'
        outstr += 'Destination account: ' + s['dest'] + '\n'
        #print("s:"+str(s))
        line_cnt = 0;
        for l in s['line']:
            outstr += '['+str(line_cnt)+'] Line Type: ' + l['line_type'] + '\n'
            if l['line_type'] == 'Invoice':
                outstr += '['+str(line_cnt)+'] Invoice Id: ' + str(l['invoice_id']) + ' Invoice Number: ' + str(l['invoice_number']) + '\n'
                outstr += '['+str(line_cnt)+'] Invoice Type: ' + l['invoice_type'] + '\n'
                outstr += '['+str(line_cnt)+'] Invoice Info: ' + l['invoice_info'] + '\n'
            elif l['line_type'] == 'Refund':
                outstr += '['+str(line_cnt)+'] Refund Id: ' + str(l['refund_id']) + '\n'
            else:
                outstr += '['+str(line_cnt)+'] Excess payment\n'
            outstr += '['+str(line_cnt)+'] Item: ' + l['i_product'] + '\n'
            outstr += '['+str(line_cnt)+'] Class: ' + l['i_class'] + '\n'
            outstr += '['+str(line_cnt)+'] Amount: ' + str(l['amount']) + '\n\n'
            line_cnt += 1
    return outstr

# dump out refunds array into a string
def dump_warefunds (refunds):
    outstr = ""
    for r in refunds:
        outstr += 'Refund Id: ' + str(r['refund_id']) + '\n'
        outstr += 'Date: ' + r['date'] + '\n'
        outstr += 'Amount: ' + str(r['amount']) + '\n'
        outstr += 'Contact: ' + r['contact_name'] + ' ('+str(r['contact_id'])+')' + '\n'
        outstr += 'Level: ' + r['level'] + '\n'
        outstr += 'Comment: ' + r['comment'] + '\n'
        outstr += 'Tender: ' + r['tender'] + '\n'
        outstr += 'Destination account: ' + r['dest'] + '\n'
        outstr += 'Payment Id: ' + str(r['payment_id']) + '\n'
        outstr += '\n'
    return outstr

def dump_contacts (contacts):
    outstr = ""
    for c in contacts:
        outstr += 'Contact id: ' + str(c.Id) + '\n'
        outstr += 'Last, First: ' + c.LastName + ', ' + c.FirstName + '\n'
        if hasattr(c, 'Status'):
            outstr += 'Status: ' + c.Status + '\n'
            outstr += 'MembershipLevel: ' + c.MembershipLevel.Name + '\n'
        else:
            outstr += 'Non-Member\n'
        outstr += '\n'
    return outstr

# return tuple of (valid, item, class) that are valid Quickbook "objects"
# for an invoice
# we don't so this for sales or donations (yet)
#
def parse_invoice (i_type, i_info, dest, amount):
    i_product = ""
    i_class = ""
    if i_type == "MembershipApplication" or i_type == "MembershipRenewal":
        i_product = QB_WA_PRODUCT_MEMBERSHIPS
        i_class = QB_WA_CLASS_MEMBERSHIPS
    elif i_type == "MembershipLevelChange":
        i_product = QB_WA_PRODUCT_MEMBERSHIPS
        i_class = QB_WA_CLASS_MEMBERSHIPS
    elif i_type == "EventRegistration":
        if i_info.find("Maker Camp") != -1:
            i_product = QB_WA_PRODUCT_SUMMER_CAMPS
            i_class = QB_WA_CLASS_SUMMER_CAMPS
        else:
            i_product = QB_WA_PRODUCT_CLASSES
            i_class = QB_WA_CLASS_CLASSES
    elif i_type == "OnlineStore":
        if i_info.find("Gift Certificate") != -1:
            i_product = QB_WA_PRODUCT_GIFTCERTIFICATE
            i_class = QB_WA_CLASS_GIFTCERTIFICATE
        else:
            return (False, i_product, i_class)
    elif dest == "Payline":
        i_product = QB_WA_PRODUCT_MEMBERSHIPS
        i_class = QB_WA_CLASS_MEMBERSHIPS
    elif i_type == "Undefined":
        i_product = QB_WA_PRODUCT_UNDEFINED
        i_class = QB_WA_CLASS_UNDEFINED
    else:
        return (False, i_product, i_class)
    return (True, i_product, i_class)

# read payments and invoices from WA
# to build a sales record to transfer to QuickBooks as a Sales Receipt
#
def build_sales_records (start, stop, debug):

    sales = []

    # fire up the API
    apiKey = get_apiKey(apiKey_fpath)
    api = WaApi.WaApiClient("CLIENT_ID", "CLIENT_SECRET")
    api.authenticate_with_apikey (apiKey, scope='account_view contacts_view finances_view')
    if debug: print('\n*** Authenticated ***')

    # Grab account details
    #
    accounts = api.execute_request("/v2/accounts")
    account = accounts[0]

    print ("Url:", account.Url)
    #for r in account.Resources:
    #    print ("Name:", r.Name, "Url:", r.Url, "Allowed:", r.AllowedOperations)

    prefix = "https://api.wildapricot.org"
    prefix_len = len(prefix)

    contactsUrl = next(res for res in account.Resources if res.Name == 'Contacts').Url
    short_contactsUrl = contactsUrl[prefix_len:]
    if debug: print('contactsUrl:', contactsUrl)

    #tendersUrl = next(res for res in account.Resources if res.Name == 'Tenders').Url
    #short_tendersUrl = tendersUrl[prefix_len:]
    #if args.debug: print('tendersUrl:', tendersUrl)

    paymentsUrl = next(res for res in account.Resources if res.Name == 'Payments').Url
    short_paymentsUrl = paymentsUrl[prefix_len:]
    if debug: print('paymentsUrl:', paymentsUrl)

    paymentallocationsUrl = next(res for res in account.Resources if res.Name == 'Payment allocations').Url
    short_paymentallocationsUrl = paymentallocationsUrl[prefix_len:]
    if debug: print('paymentallocationsUrl:', paymentallocationsUrl, 'short:', short_paymentallocationsUrl)

    invoicesUrl = next(res for res in account.Resources if res.Name == 'Invoices').Url
    short_invoicesUrl = invoicesUrl[prefix_len:]
    if debug: print('invoicesUrl:', invoicesUrl, 'short:', short_invoicesUrl)

    refundsUrl = next(res for res in account.Resources if res.Name == 'Refunds').Url
    short_refundsUrl = refundsUrl[prefix_len:]
    if debug: print('refundsUrl:', refundsUrl, 'short:', short_refundsUrl)

    # read in contacts
    contacts = get_contacts(api, debug, contactsUrl)
    chash = make_contact_hash (contacts)

    # read in payments, and build list of payment ids
    #
    payments = get_payments (api, debug, paymentsUrl, start, stop)
    payment_types = {}
    total_payments = len(payments)

    if debug: print ('Retrieved', total_payments, 'payments')

    if total_payments > 98:
        print ("Sorry, this script can't handle more than 98 payments at once.")
        print ("You have requested", total_payments, "payments. Try again with a smaller date range.")
        return ([],[],[])

    # build paymentallocations batch request from ids
    batch_req = []
    response2payment = {}
    request_cnt = 0
    payment_cnt = 0
    for p in payments:
        dest = tender2dest_payment[p.Tender.Name]
        # skip over payments that aren't invoiced (just donations as far as I know) or use
        # non-monetary tender or that haven't paid for anything
        if (p.AllocatedValue == 0) or (p.Type != 'InvoicePayment') or (dest == 'N/A'):
            print ('Skipping:')
            print (' Id:', p.Id, 'Amount:', p.Value, 'Allocated:', p.AllocatedValue, 'Refunded:', p.RefundedAmount, 'Date:', p.DocumentDate)
            print (' Contact:', p.Contact.Id, p.Contact.Name, 'Tender:', p.Tender.Id, p.Tender.Name, 'Type:', p.Type)
        else:
            payment_types[p.Type] = 1
            req = {
                "Id": "get record for payment id " + str(p.Id),
                "Order": request_cnt,
                "PathAndQuery": short_paymentallocationsUrl+"?PaymentId="+str(p.Id),
                "Method": "GET",
            }
            batch_req.append(req)
            response2payment[request_cnt] = payment_cnt
            request_cnt += 1
        payment_cnt += 1

    if debug:
        print ('Payment types found:')
        for k in payment_types:
            print ('     '+k)

    if debug: print('Making batch api call to get payment allocations')
    responses = api.execute_batch_request("https://api.wildapricot.org/batch", batch_req)

    # parse payment allocations responses and build sales records
    request_cnt = 0
    for r in responses:
        if r.HttpStatusCode != 200:
            print ('Error! HttpStatusCode:', r.HttpStatusCode)
            return ([], [], [])

        p = payments[response2payment[request_cnt]]

        level = "Non-Member"
        if hasattr (contacts[chash[p.Contact.Id]], 'MembershipLevel'):
            level = contacts[chash[p.Contact.Id]].MembershipLevel.Name

        # sales, like payments, can apply to multiple purchases (invoices)
        new_sale = {'payment_id': p.Id,
                    'date': p.DocumentDate,
                    'p_amount': p.Value,
                    'type': p.Type,
                    'contact_id': p.Contact.Id,
                    'contact_name': p.Contact.Name,
                    'level': level,
                    'tender': p.Tender.Name,
                    'dest': tender2dest_payment[p.Tender.Name],
                    'line': []}

        rda = json.loads(r.ResponseData)
        if len(rda) != 1:
            print ('multiple invoices connected to payment', p.Id, 'for', p.AllocatedValue)
        for rd in rda:
            if debug: print ("Payment:", rd['Payment']['Id'], " -> Invoice:", rd['Invoice']['Id'])
            if 'Invoice' in rd:
                new_line = {'line_type': 'Invoice',
                            'amount': rd['Value'],
                            'invoice_id': rd['Invoice']['Id']}
            else:
                # to make things work no matter if we query payments before or after a refund
                # we make up a fake "invoice" for a refund so that a refund can offset it
                new_line = {'line_type': 'Refund',
                            'amount': rd['Value'],
                            'refund_id': rd['Refund']['Id'],
                            'i_product': QB_WA_PRODUCT_REFUNDS,
                            'i_class': QB_WA_CLASS_REFUNDS}

            new_sale['line'].append(new_line)

        # add a line item for unallocated funds (excess payment) if needed
        if p.Value != p.AllocatedValue:
            print ('payment with unallocated funds found:', p.Id, 'Value:', p.Value, 'AllocatedValue:', p.AllocatedValue)
            new_line = {'line_type': 'Excess',
                        'amount': (p.Value - p.AllocatedValue),
                        'i_product': QB_WA_PRODUCT_EXCESS,
                        'i_class': QB_WA_CLASS_EXCESS}
            new_sale['line'].append(new_line)

        sales.append(new_sale)
        request_cnt += 1

    # go thru and find payments that have no allocated funds and add excess payment line items
    for p in payments:
        if (p.AllocatedValue == 0):
            level = "Non-Member"
            if hasattr (contacts[chash[p.Contact.Id]], 'MembershipLevel'):
                level = contacts[chash[p.Contact.Id]].MembershipLevel.Name
            new_sale = {'payment_id': p.Id,
                     'date': p.DocumentDate,
                     'p_amount': p.Value,
                     'type': p.Type,
                     'contact_id': p.Contact.Id,
                     'contact_name': p.Contact.Name,
                     'level': level,
                     'tender': p.Tender.Name,
                     'dest': tender2dest_payment[p.Tender.Name],
                     'line': []}
            new_line = {'line_type': 'Excess',
                        'amount': (p.Value - p.AllocatedValue),
                        'i_product': QB_WA_PRODUCT_EXCESS,
                        'i_class': QB_WA_CLASS_EXCESS}
            new_sale['line'].append(new_line)
            sales.append(new_sale)

    # build invoices batch request from payment ids
    batch_req = []
    response2sales = {}
    response2line = {}
    sales_cnt = 0
    request_cnt = 0
    for s in sales:
        line_cnt = 0
        for l in s['line']:
            if l['line_type'] == 'Invoice':
                req = {
                    "Id": "get record for invoice id " + str(l['invoice_id']),
                    "Order": request_cnt,
                    "PathAndQuery": short_invoicesUrl+str(l['invoice_id']),
                    "Method": "GET",
                }
                batch_req.append(req)
                response2sales[request_cnt] = sales_cnt
                response2line[request_cnt] = line_cnt
                request_cnt += 1
            line_cnt += 1
        sales_cnt += 1

    if debug: print('Making batch api call to get invoices')
    responses = api.execute_batch_request("https://api.wildapricot.org/batch", batch_req)
    request_cnt = 0
    for r in responses:
        if r.HttpStatusCode != 200:
            print ('Error! HttpStatusCode:', r.HttpStatusCode)
            return ([], [], [])
        rd = json.loads(r.ResponseData)
        sales_idx = response2sales[request_cnt]
        line_idx = response2line[request_cnt]
        sales[sales_idx]['line'][line_idx]['invoice_number'] = rd['DocumentNumber']
        #if sales[sales_idx]['p_amount'] != rd['PaidAmount']:
        #    print ("Warning: mismatch on amounts between Payment and Invoice")
        #    print ("Payment:", sales[sales_idx]['amount'])
        #    print ("Invoice:", rd['PaidAmount'])
        #    print ("rd:", r.ResponseData)
        i_type = rd['OrderType']
        i_info = rd['OrderDetails'][0]['Notes']
        sales[sales_idx]['line'][line_idx]['invoice_type'] = i_type
        sales[sales_idx]['line'][line_idx]['invoice_info'] = i_info
        (i_valid, i_product, i_class) = parse_invoice (i_type, i_info, sales[sales_idx]['dest'], sales[sales_idx]['line'][line_idx]['amount'])
        if i_valid:
            sales[sales_idx]['line'][line_idx]['i_product'] = i_product
            sales[sales_idx]['line'][line_idx]['i_class'] = i_class
        else:
            sales[sales_idx]['line'][line_idx]['i_product'] = 'N/A'
            sales[sales_idx]['line'][line_idx]['i_class'] = 'N/A'
        #print ('sales['+str(sales_idx)+']:', str(sales[sales_idx]))
        request_cnt += 1

    # OK, onto refunds

    wa_refunds = get_warefunds (api, debug, refundsUrl, start, stop)
    refunds = []

    # build refund batch request from refund ids
    batch_req = []
    response2refund = {}
    refund_cnt = 0
    request_cnt = 0

    for r in wa_refunds:
        level = "Non-Member"
        if hasattr (contacts[chash[r.Contact.Id]], 'MembershipLevel'):
            level = contacts[chash[r.Contact.Id]].MembershipLevel.Name
        new_refund = {'refund_id': r.Id,
                    'date': r.DocumentDate,
                    'amount': r.Value,
                    'contact_id': r.Contact.Id,
                    'contact_name': r.Contact.Name,
                    'level': level,
                    'comment': r.PublicComment,
                    'tender': r.Tender.Name,
                    'dest': tender2dest_refund[r.Tender.Name]
        }
        refunds.append(new_refund)

        req = {
            "Id": "get record for refund id " + str(r.Id),
            "Order": request_cnt,
            "PathAndQuery": short_paymentallocationsUrl+"?RefundId="+str(r.Id),
            "Method": "GET",
        }
        batch_req.append(req)
        response2refund[request_cnt] = refund_cnt
        request_cnt += 1
        refund_cnt += 1

    if debug: print('Making batch api call to get payment allocations for refunds')
    responses = api.execute_batch_request("https://api.wildapricot.org/batch", batch_req)

    # parse payment allocations responses and record payment Id
    request_cnt = 0
    for r in responses:
        if r.HttpStatusCode != 200:
            print ('Error! HttpStatusCode:', r.HttpStatusCode)
            return ([], [], [])

        rda = json.loads(r.ResponseData)
        if len(rda) != 1:
            print ('Error! multiple payments connected to refund', refunds[response2refund[request_cnt]]['refund_id'])
            return ([], [], [])
        rd = rda[0]
        if debug: print ("Refund:", rd['Refund']['Id'], " -> Payment:", rd['Payment']['Id'])
        refunds[response2refund[request_cnt]]['payment_id'] = rd['Payment']['Id']
        request_cnt += 1

    return (sales, refunds, contacts)



@app.route('/')
def index():
    html = "Please enter date range for payment records to retrieve from Wild Apricot."
    html += "<br><br>Date format is 'MM-DD-YYYY'"
    html += '<form action="/run_wa_api">'
    html += '<br>Start Date:<br><input type="text" name="start">'
    html += '<br>End Date:<br><input type="text" name="end">'
    html += '<br><input type="submit" value="Submit">'
    html += '</form>'
    return html

# <Sigh> Wild Apricot and QB use different date formats
#
def wa2qb_date (date):
    # toss everything after a "T"
    #if date.find('T') != -1:
    #    date = date[:date.find('T')]

    month, day, year = date.split('-')
    if (len(month) < 2):
        month = '0'+month
    if (len(day) < 2):
        day = '0'+day
    return year+'-'+month+'-'+day

def trim_time (date):
    # toss everything after a "T"
    if date.find('T') != -1:
        date = date[:date.find('T')]
    return date

qb_start_date = ""
qb_end_date = ""
wa_sales = []
wa_refunds = []

@app.route('/run_wa_api')
def run_wa_api():

    global qb_start_date
    global qb_end_date
    global wa_sales
    global wa_refunds

    html = ""

    wa_start_date = request.args['start']
    wa_end_date = request.args['end']

    qb_start_date = wa2qb_date (wa_start_date)
    qb_end_date = wa2qb_date (wa_end_date)

    print ('start_date:', wa_start_date, 'end_date:', wa_end_date)
    print ('start_date:', qb_start_date, 'end_date:', qb_end_date)
    (wa_sales, wa_refunds, wa_contacts) = build_sales_records (wa_start_date, wa_end_date, False)

    if len(wa_sales) == 0:
        html += "Looks like the WA API failed. Too big of a date range? Better try again."
        return html
        
    html += "WA API Success!"
    html += "<br>We retrieved "+str(len(wa_sales))+" sales records:"

    html += '<br><textarea rows="20" cols="80">'+dump_wasales(wa_sales)+'</textarea>'

    html += "<br><br>We retrieved "+str(len(wa_refunds))+" refund records:"
    if len(wa_refunds) > 0:
        html += '<br><textarea rows="10" cols="80">'+dump_warefunds(wa_refunds)+'</textarea>'

    html += "<br><br>We retrieved "+str(len(wa_contacts))+" contact records:"
    if len(wa_contacts) > 0:
        html += '<br><textarea rows="10" cols="80">'+dump_contacts(wa_contacts)+'</textarea>'

    auth_client = AuthClient(
        client_id=CLIENT_ID,
        client_secret=CLIENT_SECRET,
        environment=ENVIRONMENT,
        redirect_uri=REDIRECT_URI,
        state_token = STATE_TOKEN,
    )

    scopes = [
        Scopes.ACCOUNTING,
    ]

    # Get authorization URL
    auth_url = auth_client.get_authorization_url(scopes)

    #html = 'Main Page<br>'+'<form action="' + auth_url +'"><button type="submit">Validate</button></form>'
    html += '<br><br>'+'<a href="' + auth_url +'" target="_blank">Connect to quickbooks</a>'

    return html

auth_client = AuthClient(
    client_id=CLIENT_ID,
    client_secret=CLIENT_SECRET,
    environment=ENVIRONMENT,
    redirect_uri=REDIRECT_URI,
    state_token = STATE_TOKEN,
)

auth_code = ""
realm_id = ""
access_token = ""
refresh_token = ""

@app.route('/callback')
def callback():

    global auth_code
    global realm_id
    global access_token
    global refresh_token

    auth_code = request.args['code']
    state = request.args['state']
    realm_id = request.args['realmId']
    html = 'Callback page<br>auth_code='+auth_code+'<br>realm_id='+realm_id+'<br>state='+state+'<br>'

    if state != STATE_TOKEN:
        html = html + 'Error! state mismatch!'
    else:
        html = "Success!"
        auth_client.get_bearer_token (auth_code, realm_id=realm_id)
        access_token = auth_client.access_token
        refresh_token = auth_client.refresh_token
        #html += '<br><br><a href="http://localhost:5000/list_accounts">List Accounts</a>'
        #html += '<br><a href="http://localhost:5000/list_items">List Items</a>'
        #html += '<br><a href="http://localhost:5000/list_customers">List Customers</a>'
        #html += '<br><a href="http://localhost:5000/list_payment_methods">List Payment Methods</a>'
        #html += '<br><a href="http://localhost:5000/list_classes">List Classes</a>'
        #html += '<br><a href="http://localhost:5000/list_sales">List Sales</a>'
        #html += '<br><a href="http://localhost:5000/add_sales">Add Sales from Wild Apricot</a>'
        #html += '<br><a href="http://localhost:5000/list_refunds">List Refunds</a>'
        #html += '<br><a href="http://localhost:5000/add_refunds">Add Refunds from Wild Apricot</a>'
        html += '<br><br><a href="list_accounts">List Accounts</a>'
        html += '<br><a href="list_items">List Items</a>'
        html += '<br><a href="list_customers">List Customers</a>'
        html += '<br><a href="list_payment_methods">List Payment Methods</a>'
        html += '<br><a href="list_classes">List Classes</a>'
        html += '<br><a href="list_sales">List Sales</a>'
        html += '<br><a href="add_sales">Add Sales from Wild Apricot</a>'
        html += '<br><a href="list_refunds">List Refunds</a>'
        html += '<br><a href="add_refunds">Add Refunds from Wild Apricot</a>'

    return html

def get_customers(client):
    customers = Customer.all(max_results=500, qb=client)
    return customers

def find_customer(customers, name):
    for c in customers:
        if c.DisplayName == name:
            return c.Id
    return ""

def add_customer (customers, name, client):
    if find_customer (customers, name) != "":
        return

    parent_ref = find_customer (customers, "Wild Apricot")
    if parent_ref == "":
        print ("Error: Wild Apricot customer not found")
        return

    customer = Customer()
    customer.DisplayName = name
    customer.PreferredDeliveryMethod = "Print"

    pref = Ref()
    pref.name = "Wild Apricot"
    pref.value = str(parent_ref)
    customer.ParentRef = pref

    customer.Job = True
    customer.BillWithParent = False

    customer.save (qb=client)

def print_customer(c):
    html = 'DisplayName: ' + c.DisplayName + '\n'
    html += 'Id: '+ c.Id + '\n'
    html += 'FullyQualifiedName: ' + c.FullyQualifiedName + '\n'
    html += '\n'
    return html

@app.route('/list_customers')
def list_customers():

    html = 'Customers:'
    html += '<br><textarea rows="20" cols="80">'
    client = QuickBooks ( auth_client=auth_client, refresh_token=refresh_token, company_id=realm_id)
    customers = get_customers (client)
    for c in customers:
        html += print_customer (c)

    html += '</textarea>'

    return html

def get_accounts (client):
    #accts = Account.all(max_results=500, qb=client)
    accts = Account.where("Active = True", max_results=500, qb=client)
    return (accts)

def find_account(accounts, name):
    for a in accounts:
        if a.Name == name:
            return a.Id
    return ''

@app.route('/list_accounts')
def list_accounts():
    html = ""
    client = QuickBooks ( auth_client=auth_client, refresh_token=refresh_token, company_id=realm_id)
    accts = get_accounts (client)
    html += '<br><textarea rows="20" cols="80">'
    for a in accts:
        html += 'Name: ' + a.Name + '\n'
        html += 'Number: ' + a.AcctNum + '\n'
        html += 'Full Name: ' + a.FullyQualifiedName + '\n'
        html += 'Id: ' + a.Id + '\n'
        html += '\n'
    html += '</textarea>'
    return html

def get_items(client):
    items = Item.all(max_results=500, qb=client)
    return items

def find_item(items, name):
    for i in items:
        if i.Name == name:
            return i.Id
    return ''

@app.route('/list_items')
def list_items():
    html = ""
    html += '<br><textarea rows="20" cols="80">'
    client = QuickBooks ( auth_client=auth_client, refresh_token=refresh_token, company_id=realm_id)
    items = get_items(client)
    for i in items:
        html += 'Name: ' + i.Name + ' id: ' + i.Id + '\n'
        html += 'FullyQualifiedName: ' + i.FullyQualifiedName + '\n'
        html += 'Type: ' + i.Type + '\n'
        html += 'Income Account: ' + i.IncomeAccountRef.name + ' (' + i.IncomeAccountRef.value + ')' + '\n\n'
    html += '</textarea>'
    return html

def get_pmethods (client):
    pmethods = PaymentMethod.all(max_results=500, qb=client)
    return pmethods

def find_pmethod(pmethods, name):
    for pm in pmethods:
        if pm.Name == name:
            return pm.Id
    return ''

@app.route('/list_payment_methods')
def list_payment_methods():
    html = ""
    html += '<br><textarea rows="20" cols="80">'
    client = QuickBooks ( auth_client=auth_client, refresh_token=refresh_token, company_id=realm_id)
    pmethods = get_pmethods(client)
    for pm in pmethods:
        html += 'Name: ' + pm.Name + ' id: ' + pm.Id + '\n'
        html += 'Type: ' + pm.Type + '\n\n'
    html += '</textarea>'
    return html

def get_classes (client):
    classes = Class.all(max_results=500, qb=client)
    return classes

def find_class(classes, name):
    for c in classes:
        if c.Name == name:
            return c.Id
    return ''

@app.route('/list_classes')
def list_classes():
    html = ""
    html += '<br><textarea rows="20" cols="80">'
    client = QuickBooks ( auth_client=auth_client, refresh_token=refresh_token, company_id=realm_id)
    classes = get_classes(client)
    for c in classes:
        html += 'Name: ' + c.Name + ' id: ' + c.Id + '\n'
        html += 'FullyQualifiedName: ' + c.FullyQualifiedName + '\n\n'
    html += '</textarea>'
    return html

def get_sales(start_date, end_date, client):
    #sales = SalesReceipt.all(max_results=100, qb=client)
    query =  "TxnDate >= '"+qb_start_date+"' AND TxnDate <= '"+qb_end_date+"'"
    print ('query = "'+query+'"')
    sales = SalesReceipt.where(query, qb=client)
    #sales = SalesReceipt.query("SELECT * FROM SalesReceipt WHERE TxnDate = '2019-07-01'", qb=client)
    return sales

def get_single_sale (doc, client):
    sales = SalesReceipt.where("DocNumber = '"+str(doc)+"'", qb=client)
    return sales[0]

def find_sale(sales, doc):
    for s in sales:
        if s.DocNumber == doc:
            return s.Id
    return ""


@app.route('/list_sales')
def list_sales():

    html = ""
    client = QuickBooks ( auth_client=auth_client, refresh_token=refresh_token, company_id=realm_id)

    sales = get_sales (qb_start_date, qb_end_date, client)

    if len(sales) > 0:
        html += 'Sales:'
        html += '<br><textarea rows="20" cols="80">'
        for s in sales:
            html += 'Date: ' + s.TxnDate + ' Id: '+ s.Id + '\n'
            html += 'DocNumber: ' + str(s.DocNumber) + ' Id: '+ s.Id + '\n'
            html += 'Customer: ' + s.CustomerRef.name + ' ('+ s.CustomerRef.value + ')' + '\n'
            html += 'Deposit Account: ' + s.DepositToAccountRef.name + ' ('+ s.DepositToAccountRef.value + ')' + '\n'
            html += 'Total Amount: ' + str(s.TotalAmt) + '\n'
            for l in s.Line:
                if l.DetailType == "SalesItemLineDetail":
                    html += '    Descripton: ' + l.Description + '\n'
                    html += '    Amount: ' + str(l.Amount) + '\n'
                    html += '    Qty: ' + str(l.SalesItemLineDetail['Qty']) + ' Unit Price: ' + str(l.SalesItemLineDetail['UnitPrice']) + '\n'
                    html += '    Item: ' + l.SalesItemLineDetail["ItemRef"]["name"] + ' (' + l.SalesItemLineDetail["ItemRef"]["value"] + ')' + '\n'
                    html += '    Class: ' + l.SalesItemLineDetail["ClassRef"]["name"] + ' (' + l.SalesItemLineDetail["ClassRef"]["value"] + ')' + '\n'
            #html += 'Raw:'
            #html += str(s.to_json())
            html += '\n'
        html += '</textarea>'
    else:
        html += 'No sales in date range'

    return html


@app.route('/add_sales')
def add_sales():
    html = ""
    client = QuickBooks ( auth_client=auth_client, refresh_token=refresh_token, company_id=realm_id)

    # get accounts
    accts = get_accounts (client)

    # get customers
    customers = get_customers (client)

    # get payment methods
    pmethods = get_pmethods (client)

    # get items
    items = get_items (client)

    # get classes
    classes = get_classes (client)

    # get existing sales over time period so we don't duplicate
    sales = get_sales (qb_start_date, qb_end_date, client)

    failed = False
    missing_accts = {}
    missing_customers = {}
    missing_pmethods = {}
    missing_items = {}
    missing_classes = {}
    # for each sale from wild apricot, make sure everything can proceed
    for was in wa_sales:
        # make sure account, customer, payment method, item and class have been created for the sales
        dep_ref = find_account (accts, was['dest'])
        if dep_ref == "":
            failed = True
            missing_accts[was['dest']] = 1
        else:
            was['dep_ref'] = dep_ref

        cust_ref = find_customer (customers, was['level'])
        if cust_ref == "":
            failed = True
            missing_customers[was['level']] = 1
        else:
            was['cust_ref'] = cust_ref

        pmethod_ref = find_pmethod (pmethods, was['tender'])
        if pmethod_ref == "":
            failed = True
            missing_pmethods[was['tender']] = 1
        else:
            was['pmethod_ref'] = pmethod_ref

        for wal in was['line']:

            item_ref = find_item (items, wal['i_product'])
            if item_ref == "":
                failed = True
                missing_items[wal['i_product']] = 1
            else:
                wal['item_ref'] = item_ref

            class_ref = find_class (classes, wal['i_class'])
            if class_ref == "":
                failed = True
                missing_classes[wal['i_class']] = 1
            else:
                wal['class_ref'] = class_ref

    if failed:
        html += '<br>Failure!'
        html += '<br>Missing accounts:'
        for a in missing_accts:
            html += '<br>    '+a
        html += '<br>Missing customers (automatically added):'
        for c in missing_customers:
            html += '<br>    '+c
            add_customer (customers, c, client)
        html += '<br>Missing payment methods:'
        for pm in missing_pmethods:
            html += '<br>    '+pm
        html += '<br>Missing items:'
        for i in missing_items:
            html += '<br>    '+i
        html += '<br>Missing Classes:'
        for c in missing_classes:
            html += '<br>    '+c
        return html

    # for each sale from wild apricot, create a sales receipt in quickbooks
    for was in wa_sales:

        docnumber = str(was['payment_id'])
        existing = find_sale (sales, docnumber)
        if existing != "":
            print ("Skipping existing sales receipt", docnumber)
            continue

        sr = SalesReceipt()

        # we are using aggregated customer names based on membership levels

        custref = Ref()
        custref.name = was['level']
        custref.value = str(was['cust_ref'])
        sr.CustomerRef = custref

        txndate = trim_time (was['date'])
        print ('Sales txndate:', txndate)
        sr.TxnDate = txndate
        sr.DocNumber = docnumber
        sr.PrivateNote = was['contact_name'] + ' (' + str(was['contact_id']) + ')'

        depref = Ref()
        depref.name = was['dest']
        depref.value = str(was['dep_ref'])
        sr.DepositToAccountRef = depref

        pmref = Ref()
        pmref.name = was['tender']
        pmref.value = str(was['pmethod_ref'])
        sr.PaymentMethodRef = pmref


        for wal in was['line']:
            if wal['i_product'] == 'N/A':
                print ('Skipping over line in payment')
                continue
            line = SalesItemLine()
            line.Description = wal['line_type']
            line.Amount = wal['amount']

            detail = SalesItemLineDetail()
            detail.Qty = 1
            detail.UnitPrice = wal['amount']

            iref = Ref()
            iref.name = wal['i_product']
            iref.value = str(wal['item_ref'])
            detail.ItemRef = iref

            cref = Ref()
            cref.name = wal['i_class']
            cref.value = str(wal['class_ref'])
            detail.ClassRef = cref

            detail.ServiceDate = txndate

            line.SalesItemLineDetail = detail

            sr.Line.append(line)

        sr.save(qb=client)
        html += '<br>Created salesreceipt '+sr.Id+ ' from payment '+docnumber

    return html

def get_refunds(start_date, end_date, client):
    query =  "TxnDate >= '"+qb_start_date+"' AND TxnDate <= '"+qb_end_date+"'"
    print ('query = "'+query+'"')
    refunds = RefundReceipt.where(query, qb=client)
    return refunds

def find_refund(refunds, doc):
    for r in refunds:
        if r.DocNumber == doc:
            return r.Id
    return ""

def dump_refund(r):
    html = ""
    html += 'Date: "' + r.TxnDate + '" Id: "'+ str(r.Id) + '"\n'
    html += 'DocNumber: "' + str(r.DocNumber) + '"\n'
    html += 'Customer: "' + r.CustomerRef.name + '" ("'+ r.CustomerRef.value + '")' + '\n'
    html += 'Private Note: "' + str(r.PrivateNote) + '"\n'
    html += 'Deposit Account: "' + r.DepositToAccountRef.name + '" ("'+ r.DepositToAccountRef.value + '")' + '\n'
    html += 'Payment Method: "' + r.PaymentMethodRef.name + '" ("'+ r.PaymentMethodRef.value + '")' + '\n'
    #html += 'Payment Ref Num: "' + r.PaymentRefNum + '"\n'
    html += 'Print Status: "' + r.PrintStatus + '"\n'
    #html += 'Check Number: "' + r.CheckPayment + '"\n'
    html += 'Total Amount: ' + str(r.TotalAmt) + ' (type = ' + str(type(r.TotalAmt)) + ')\n'
    for l in r.Line:
        if l.DetailType == "SalesItemLineDetail":
            html += '    Descripton: "' + l.Description + '"\n'
            html += '    Amount: ' + str(l.Amount) + ' (type = ' + str(type(l.Amount)) + ')\n'
            html += '    Qty: ' + str(l.SalesItemLineDetail['Qty']) + ' (type = ' + str(type(l.SalesItemLineDetail['Qty'])) + ') Unit Price: ' + str(l.SalesItemLineDetail['UnitPrice']) + ' (type = ' + str(type(l.SalesItemLineDetail['UnitPrice'])) + ')\n'
            html += '    Item: "' + l.SalesItemLineDetail["ItemRef"]["name"] + '" ("' + l.SalesItemLineDetail["ItemRef"]["value"] + '")' + '\n'
            html += '    Class: "' + l.SalesItemLineDetail["ClassRef"]["name"] + '" ("' + l.SalesItemLineDetail["ClassRef"]["value"] + '")' + '\n'
            #html += '    ServiceDate: "' + l.SalesItemLineDetail["ServiceDate"] + '"\n'
            html += '\n'

    #html += "raw:\n"
    #html += str(r.to_json())
    return html

def dump_refund_to_add(r):
    html = ""
    html += 'Date: "' + r.TxnDate + '" Id: "'+ str(r.Id) + '"\n'
    html += 'DocNumber: "' + str(r.DocNumber) + '"\n'
    html += 'Customer: "' + r.CustomerRef.name + '" ("'+ r.CustomerRef.value + '")' + '\n'
    html += 'Private Note: "' + str(r.PrivateNote) + '"\n'
    html += 'Deposit Account: "' + r.DepositToAccountRef.name + '" ("'+ r.DepositToAccountRef.value + '")' + '\n'
    html += 'Payment Method: "' + r.PaymentMethodRef.name + '" ("'+ r.PaymentMethodRef.value + '")' + '\n'
    #html += 'Payment Ref Num: "' + r.PaymentRefNum + '"\n'
    html += 'Print Status: "' + r.PrintStatus + '"\n'
    #html += 'Check Number: "' + r.CheckPayment + '"\n'
    #html += 'Total Amount: ' + str(r.TotalAmt) + ' (type = ' + str(type(r.TotalAmt)) + ')\n'
    for l in r.Line:
        if l.DetailType == "SalesItemLineDetail":
            html += '    Descripton: "' + l.Description + '"\n'
            html += '    Amount: ' + str(l.Amount) + ' (type = ' + str(type(l.Amount)) + ')\n'
            html += '    Qty: ' + str(l.SalesItemLineDetail.Qty) + ' (type = ' + str(type(l.SalesItemLineDetail.Qty)) + ') Unit Price: ' + str(l.SalesItemLineDetail.UnitPrice) + ' (type = ' + str(type(l.SalesItemLineDetail.UnitPrice)) + ')\n'
            html += '    Item: "' + l.SalesItemLineDetail.ItemRef.name + '" ("' + l.SalesItemLineDetail.ItemRef.value + '")' + '\n'
            html += '    Class: "' + l.SalesItemLineDetail.ClassRef.name + '" ("' + l.SalesItemLineDetail.ClassRef.value + '")' + '\n'
            #html += '    ServiceDate: "' + l.SalesItemLineDetail.ServiceDate + '"\n'
            html += '\n'
    #html += "raw:\n"
    #html += str(r.to_json())
    return html



@app.route('/list_refunds')
def list_refunds():

    html = ""
    client = QuickBooks ( auth_client=auth_client, refresh_token=refresh_token, company_id=realm_id)

    refunds = get_refunds (qb_start_date, qb_end_date, client)

    if len(refunds) > 0:
        html += 'Refunds:'
        html += '<br><textarea rows="20" cols="80">'
        for r in refunds:
            html += dump_refund(r)
        html += '</textarea>'
    else:
        html += 'No refunds in date range'

    return html

# given a refund, find the line from the sale object that matches
def find_matching_line (sale, wa_refund):
    
    ref_amount = wa_refund['amount']

    # build a list of matching amounts
    price_matches = []
    for l in sale.Line:
        if l.DetailType == "SalesItemLineDetail":
            if l.Amount == -ref_amount:
                price_matches.append(l)

    if len(price_matches) == 0:
        print ("Error! Couldn't find matching item for refund "+str(wa_refund['refund_id'])+" in sale/receipt "+str(sale.DocNumber))
        return False

    # go through that list of lines and return the first refund item, else return the first line
    for l in price_matches:
        if l.SalesItemLineDetail['ItemRef']['name'] == 'Refund':
            return (l)

    return (price_matches[0])


@app.route('/add_refunds')
def add_refunds():
    html = ""
    client = QuickBooks ( auth_client=auth_client, refresh_token=refresh_token, company_id=realm_id)

    # get accounts
    accts = get_accounts (client)

    # get customers
    customers = get_customers (client)

    # get payment methods
    pmethods = get_pmethods (client)

    # get items
    items = get_items (client)

    # get classes
    classes = get_classes (client)

    # get existing sales over time period so we don't duplicate
    refunds = get_refunds (qb_start_date, qb_end_date, client)

    failed = False
    missing_accts = {}
    missing_pmethods = {}
    missing_customers = {}
    # for each refund from wild apricot, make sure everything can proceed
    for war in wa_refunds:
        # make sure account, payment method and salesreceipt exist
        dep_ref = find_account (accts, war['dest'])
        if dep_ref == "":
            failed = True
            missing_accts[war['dest']] = 1
        else:
            war['dep_ref'] = dep_ref

        cust_ref = find_customer (customers, war['level'])
        if cust_ref == "":
            failed = True
            missing_customers[war['level']] = 1
        else:
            war['cust_ref'] = cust_ref

        pmethod_ref = find_pmethod (pmethods, war['tender'])
        if pmethod_ref == "":
            failed = True
            missing_pmethods[war['tender']] = 1
        else:
            war['pmethod_ref'] = pmethod_ref



    if failed:
        html += '<br>Failure!'
        html += '<br>Missing accounts:'
        for a in missing_accts:
            html += '<br>    '+a
        html += '<br>Missing cutomers:'
        for a in missing_customers:
            html += '<br>    '+a
        html += '<br>Missing payment methods:'
        for pm in missing_pmethods:
            html += '<br>    '+pm
        return html

    html += 'Refunds Added:'
    html += '<br><textarea rows="20" cols="80">'

    # for each refund from wild apricot, create a sales receipt in quickbooks
    for war in wa_refunds:

        docnumber = str(war['refund_id'])
        existing = find_refund (refunds, docnumber)
        if existing != "":
            print ("Skipping existing refund", docnumber)
            continue

        # lookup up corresponding sale...
        sale = get_single_sale (war['payment_id'], client)
        
        # ...and find matching line item
        sales_line = find_matching_line (sale, war)
        if sales_line == False:
            # in theory, this should never happen
            continue

        ref = RefundReceipt()

        # we are using aggregated customer names based on membership levels

        custref = Ref()
        custref.name = war['level']
        custref.value = str(war['cust_ref'])
        ref.CustomerRef = custref

        txndate = trim_time (war['date'])
        print ('Refund txndate:', txndate)
        ref.TxnDate = txndate
        ref.DocNumber = str(docnumber)
        ref.PrivateNote = war['contact_name'] + ' (' + str(war['contact_id']) + ')'

        depref = Ref()
        depref.name = war['dest']
        depref.value = str(war['dep_ref'])
        ref.DepositToAccountRef = depref

        pmref = Ref()
        pmref.name = war['tender']
        pmref.value = str(war['pmethod_ref'])
        ref.PaymentMethodRef = pmref

        ref.PrintStatus = "NeedToPrint"

        line = SalesItemLine()
        line.Description = "Refund"
        line.Amount = -war['amount']
        line.LineNum = 1

        detail = SalesItemLineDetail()
        detail.Qty = 1
        detail.UnitPrice = -war['amount']

        iref = Ref()
        iref.name = sales_line.SalesItemLineDetail['ItemRef']['name']
        iref.value = str(sales_line.SalesItemLineDetail['ItemRef']['value'])
        detail.ItemRef = iref

        cref = Ref()
        cref.name = sales_line.SalesItemLineDetail['ClassRef']['name']
        cref.value = str(sales_line.SalesItemLineDetail['ClassRef']['value'])
        detail.ClassRef = cref

        detail.ServiceDate = txndate

        line.SalesItemLineDetail = detail

        ref.Line.append(line)

        # these fields cause api to fail
        del ref.CustomField
        del ref.CustomerMemo
        del ref.CheckPayment
        del ref.CreditCardPayment
        del ref.PaymentType

        ref.save(qb=client)
        html += 'Created QB Refund '+str(ref.Id)+ ' from Wild Apricot refund '+docnumber+'\n'
        html += dump_refund_to_add (ref)

    html += '</textarea>'
    return html
