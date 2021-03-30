#!/usr/bin/python
# -*- coding: utf-8 -*-

import os
import sqlite3
import xlwt
import re
import uuid
from xlrd import open_workbook
from datetime import datetime
from nameparser import HumanName


def regexp(expr, item):
    reg = re.compile(expr)
    return reg.search(item) is not None


def get_cursor(SQLITE_FILE):
    # Connecting to the database file
    CONN = sqlite3.connect(SQLITE_FILE)
    CONN.create_function("REGEXP", 2, regexp)

    # create cursor
    cursor = CONN.cursor()
    cursor.execute("PRAGMA synchronous = OFF")
    cursor.execute("PRAGMA journal_mode=OFF")
    cursor.execute("PRAGMA locking_mode = EXCLUSIVE")
    cursor.execute("PRAGMA temp_store = MEMORY")
    cursor.execute("PRAGMA count_changes = OFF")
    cursor.execute("PRAGMA PAGE_SIZE = 4096")

    return cursor


ClientId = 125083
FileId = "eeea5607-5a2c-47e4-8263-498b2d2d9d8d"
PriorProviderId = ""
FileVersion = ""
FileName = ""
FileUploadDateTime = ""
FileUploadedByUser = ""


def get_import_model():
    import_model = ImportModel()
    import_model.clientId = ClientId
    import_model.fileId = FileId
    import_model.fileName = FileName
    import_model.fileUploadDateTime = FileUploadDateTime
    import_model.fileUploadedByUser = FileUploadedByUser
    # import_model.fileSequenceId = str(uuid.uuid4())
    import_model.providerId = PriorProviderId
    return import_model


def create_import_model(clientId, fileId, fileName, fileUploadDateTime, fileUploadedByUser, priorProviderId):
    import_model = ImportModel()
    import_model.clientId = clientId
    import_model.fileId = fileId
    import_model.fileName = fileName
    import_model.fileUploadDateTime = fileUploadDateTime
    import_model.fileUploadedByUser = fileUploadedByUser
    import_model.fileSequenceId = str(uuid.uuid4())
    import_model.providerId = priorProviderId
    return import_model


# lookup definitions
def get_lookup_value(value, lookups):
    if value is None:
        return None
    upper_lookups = {k.upper().strip(): v for k, v in lookups.items()}

    if value.upper().strip() in upper_lookups:
        return upper_lookups[value.upper().strip()]

    if len(value) == 0:
        return None

    return value.strip()


yesNoLookup = {
    "Yes": "Y",
    "No": "N"
}

prefixLookup = {
    "Mr.": "0000",
    "Ms.": "0001",
    "Miss": "0002",
    "Mrs.": "0003",
    "Rev.": "0004",
    "Sr.": "0005",
    "Fr.": "0006",
    "Dr. ": "0007",
    "Prof.": "0008",
}

suffixLookup = {
    "Jr.": "0000",
    "Sr.": "0001",
    "II": "0002",
    "III": "0003",
    "IV": "0004",
    "V": "0005",
    "VI": "0006",
    "VII": "0007",
    "VIII": "0008",
    "IX": "0009",
    "X": "0010",
}

reciprocityLookup = {
    "Live In": "L",
    "Work In": "W"
}

frequencyLookup = {
    "Every pay period": "0",
    "First pay period of month": "1",
    "Second pay period of mnth": "2",
    "Third pay period of month": "3",
    "Fourth pay period of mnth": "4",
    "Every / not 3rd / not 5th": "5",
    "1st and 3rd pay period": "6",
    "2nd and 4th pay period": "7",
    "Last pay period": "8",
    "Special Occurrence": "9",
    "1st and 2nd pay period": "A",
}

adpEthnicity = {
    '1': 'White',
    '2': 'Black or African American',
    '3': 'Hispanic or Latino',
    '4': 'Asian',
    '5': 'Amer. Ind or AK Native',
    '6': 'Nat HI or Oth Pac Island',
    '9': 'Two or more Races',
    '': ''
}

ethnicityLookup = {
    "Amer. Ind or AK Native": "1",
    "Asian": "2",
    "Black or African American": "3",
    "Hispanic or Latino": "5",
    "White": "6",
    "Nat HI or Oth Pac Island": "7",
    "Two or more Races": "8",
    "Declined to Identify": "9",
    "Not Hispanic or Latino": "",
    "W": "6",
    "B": "3",
    "H": "5",
    "A": "2",
    "T": "8",
    "U": "7",
}

adjustWithholdingLookup = {
    "Add to withholding": "addTo",
    "Override withholding": "Block",
    "None": "Use",
}

directDepositLookup = {
    "Net": "N",
    "Partial": "P",
    "Amount": "D",
    'Rate': 'P'
}

accountLookup = {
    "Checking": "C",
    "Savings": "S",
}

filingStatusType = {
    "Married": "M",
    "Single": "S",
    "All / A-CT": "A",
    "B": "B",
    "C": "C",
    "D": "D",
    "F": "F",
    "Head of Household": "H",
    "Married Filing Jointly": "J",
    "Joint - 2 incomes": "J2",
    "Married/Joint-KS": "M",
    "Married or Joint-KS": "M",
    "Married - 2 incomes": "M2",
    "Married Filing Separately": "MS",
    "Married, but withhold at higher Single rate": "S",
    "No Personal Exemptions": "O",
    "Single or Head of Househo": "H",
    "Single or Head of Househould": "H",
    "Spouse also filing W5": "M",
    "No Spouse/Spouse no W5": "S",
}

employeeType = {
    "Casual": "C",
    "Independent Contractor": "I",
    "Regular": "R",
    "RFT": "R",
    "RPT": "R",
    "Seasonal": "S",
    "SFT": "S",
    "SPT": "S",
    "Temporary": "T",
    "TFT": "T",
    "TPT": "T",
    "TEMP": "T",
    "Variable": "V",
}

statusLookup = {
    "Full Time": "Y",
    "F": "Y",
    "RFT": "Y",
    "SFT": "Y",
    "TFT": "Y",
    "EFT": "Y",
    "FT": "Y",
    "Part Time": "N",
    "P": "N",
    "RPT": "N",
    "SPT": "N",
    "TPT": "N",
    "PT": "N",
    "PART": "N",
}

gender = {
    "Female": "F",
    "Male": "M",
    "F": "F",
    "M": "M",
}

maritalStatusLookup = {
    "Married": "M",
    "Single": "S",
    "Divorced": "D",
    "State-Recognized Union": "L",
    "Widowed": "W",
    "Separated": "X",
    "M": "M",
    "S": "S",
    "D": "D",
    "W": "W",
}

employmentStatus = {
    "Active": "A",
    "Deceased": "D",
    "Disability - long": "DL",
    "Disability - short": "DS",
    "FMLA": "FL",
    "Laid off": "LO",
    "Leave with pay": "LP",
    "Leave without pay": "LW",
    "3rd-prty payable": "P",
    "Resigned": "R",
    "Retired": "RT",
    "Terminated": "T",
    "Wkrs Comp": "WC",
    "A": "A",
    "T": "T",
    "L": "LW"
}

flsaLookup = {
    "Hourly Exempt": "E",
    "Hourly Non Exempt": "N",
    "Salary Exempt": "P",
    "Salary Non Exempt": "S"
}

# dictionary of special characters found in phone numbers
# 'i':'j' = 'find':'replace'
phone_specialCharacters = {
    '(': '',
    ')': '',
    '-': '',
    ' ': ''
}

suffix_dictonary = {
    'Jr': 'Jr.',
    'Jr.': 'Jr.',
    'JR': 'Jr.',
    'JR.': 'Jr.',
    'Junior': 'Jr.',
    'Sr': 'Sr.',
    'Sr.': 'Sr.',
    'SR': 'Sr.',
    'SR.': 'Sr.',
    'Senior': 'Sr.',
    'II': 'II',
    'Ii': 'II',
    'iI': 'II',
    'ii': 'II',
    'III': 'III',
    'Iii': 'III',
    'iii': 'III',
    'IV': 'IV',
    'Iv': 'IV',
    'iv': 'IV',
    'v': 'V',
    'V': 'V',
    'vi': 'VI',
    'Vi': 'VI',
    'VI': 'VI',
    'the second': 'II',
    'The Second': 'II',
    'The second': 'II',
    'THE SECOND': 'II',
    'the third': 'III',
    'The Third': 'III',
    'The third': 'III',
    'THE THIRD': 'III',
    'the fourth': 'IV',
    'The Fourth': 'IV',
    'The fourth': 'IV',
    'THE FOURTH': 'IV',
    'the fifth': 'V',
    'The Fifth': 'V',
    'The fifth': 'V',
    'THE FIFTH': 'V'
}

payFreq_dictonary = {
    'Bi-weekly': 'Bi-Weekly',
}

us_states = ["AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA", "HI", "ID", "IL", "IN", "IA", "KS", "KY", "LA",
             "ME", "MD", "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ", "NM", "NY", "NC", "ND", "OH", "OK",
             "OR", "PA", "RI", "SC", "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV", "WI", "WY"]

state_tax_codes = ["AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA", "HI", "ID", "IL", "IN", "IA", "KS", "KY",
                   "LA", "ME", "MD", "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ", "NM", "NY", "NC", "ND",
                   "OH", "OK", "OR", "PA", "RI", "SC", "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV", "WI", "WY", "AS",
                   "DC", "FM", "GU", "MH", "MP", "PW", "PR", "VI"]

id_headers = ['priorEmployeeNumber', 'lastName', 'firstName', 'birthDate', 'ssn']

employee_headers = ['action', 'clientId', 'priorCompanyCode', 'priorEmployeeNumber', 'employeeNumber', 'ssn',
                    'birthDate', 'lastName', 'firstName', 'middleName', 'prefix', 'suffix', 'accredited',
                    'addressLine1', 'addressLine2', 'city', 'state', 'zip', 'departmentCode', 'departmentDescription',
                    'payrollCode', 'paygroupDescription', 'employmentStatus', 'terminationDate', 'terminationReason',
                    'reHireDate', 'hireDate', 'flsa', 'statusType', 'employeeType', 'maritalStatus', 'gender',
                    'ethnicity', 'jobTitle', 'workPhone', 'workPhoneNumberExtension', 'workEmail', 'mobilePhone',
                    'homePhone', 'homeEmail', 'annualHours', 'ownerOfficer', 'baseShift', 'adjustedHireDate',
                    'seniorityDate', 'checkPrintSort', 'managerEmployeeNumber', 'managerFirstName', 'managerLastName',
                    'managerClientId']

tax_headers = ['action', 'clientId', 'priorCompanyCode', 'priorEmployeeNumber', 'ssn', 'birthDate', 'lastName',
               'firstName', 'taxCode', 'altCodes', 'filingStatus', 'exemptions', 'adjustWithHolding', 'percentage',
               'amount', 'reciprocity', 'blockDate', 'calculate', 'spouseWork', 'additionalStateExemptions',
               'nonResidentAlienAdditionalFitwh', 'ncciCode', 'psdCode', 'psdRate', 'blockIndicator',
               'employmentStatus', 'terminationDate']

earning_headers = ['action', 'clientId', 'priorCompanyCode', 'priorEmployeeNumber', 'ssn', 'birthDate', 'lastName',
                   'firstName', 'effectiveDates effectiveDate1', 'code', 'altCodes', 'effectiveDates amount1',
                   'effectiveDates rate1', 'calculate', 'policyAmount', 'hours', 'employmentStatus', 'terminationDate']

deduction_headers = ['action', 'clientId', 'priorCompanyCode', 'priorEmployeeNumber', 'ssn', 'birthDate', 'lastName',
                     'firstName', 'effectiveDates effectiveDate1', 'code', 'altCodes', 'effectiveDates rate1',
                     'effectiveDates amount1', 'calculate', 'limits maxAmount1', 'AmtToDate', 'frequency', 'deductionType',
                     'employmentStatus', 'terminationDate']

payrate_headers = ['action', 'clientId', 'priorCompanyCode', 'priorEmployeeNumber', 'ssn', 'birthDate', 'lastName',
                   'firstName', 'effectiveDate', 'sequence', 'payType', 'payRate', 'description', 'reason',
                   'employmentStatus', 'terminationDate']

directdeposit_headers = ['action', 'clientId', 'priorCompanyCode', 'priorEmployeeNumber', 'ssn', 'birthDate',
                         'lastName', 'firstName', 'deductionCode', 'routingNumber', 'accountNumber', 'accountType',
                         'rate', 'amount', 'calculate', 'directDepositType', 'frequency', 'employmentStatus',
                         'terminationDate']

parallel_headers = ['priorEmployeeId', 'firstName', 'lastName', 'checkDate', 'checkNumber', 'priorDeptId', 'itemType',
                    'itemPriorCode', 'amount', 'hours', 'rate']

lastNameMultipartStrings = ['DE', 'DA', 'DI', 'VON', 'VAN', 'LE', 'LA', 'DU', 'DES', 'DEL', 'DE LA', 'DELLA', 'VAN DER',
                            'ST', 'SAINT']


# class definitions
class Account:
    account = None
    accountType = None
    amount = None
    code = None
    altCodes = None
    routing = None
    active = None
    percent = None
    frequency = None
    type = None
    deductionCode = None

    def __init__(self):
        self.account = None
        self.accountType = None
        self.amount = None
        self.code = None
        self.altCodes = None
        self.routing = None
        self.active = None
        self.percent = None
        self.frequency = None
        self.type = None
        self.deductionCode = None


class Pay:
    payRates = []

    def __init__(self):
        self.payRates = []


class Names:
    first = None
    middle = None
    last = None
    prefix = None
    suffix = None

    def __init__(self):
        self.first = None
        self.middle = None
        self.last = None
        self.prefix = None
        self.suffix = None


class Manager:
    firstName = None
    lastName = None
    employeeId = None

    def __init__(self):
        self.firstName = None
        self.lastName = None
        self.employeeId = None


class ClientDepartment:
    id = None
    name = None

    def __init__(self):
        self.id = None
        self.name = None

    # hash, eq, and ne necessary for using Set againts an object, this
    # classes properties should be treated as immutable

    def __hash__(self):
        return hash((self.id, self.name))

    def __eq__(self, other):
        return ((self.id == other.id) and (self.name == other.name))

    def __ne__(self, other):
        return (not self.__eq__(other))


class ImportModel:
    providerId = None
    clientId = None
    fileId = None
    versionNumber = None
    fileName = None
    fileUploadDateTime = None
    fileUploadedByUser = None
    employees = []
    earningCodes = []
    taxCodes = []
    deductionCodes = []
    departments = []

    def __init__(self):
        self.providerId = None
        self.clientId = None
        self.fileId = None
        self.versionNumber = None
        self.fileName = None
        self.fileUploadDateTime = None
        self.fileUploadedByUser = None
        self.employees = []
        self.earningCodes = []
        self.taxCodes = []
        self.deductionCodes = []
        self.departments = []


class ClientCode:
    code = None
    altCodes = None
    amount = None
    rate = None
    effectiveDate = None
    calculate = None
    limit = None

    def __init__(self):
        self.code = None
        self.altCodes = None
        self.amount = None
        self.rate = None
        self.effectiveDate = None
        self.calculate = None
        self.limit = None


class TaxCode:
    code = None
    amount = None
    rate = None
    exemptions = None
    exemptionAmount = None
    exemptionPercent = None
    block = False
    type = None
    filingStatus = None
    headOfHousehold = False
    psdRate = None
    psdCode = None
    nonResidentAlienAdditionalFitwh = None
    applicableBirthYear = None
    adjustWithholding = None
    spouseWork = None
    livedInWorkedIn = None
    calculate = None
    numberOfDependents = None
    numberOfOtherDependents = None
    hasTwoIncomes = False
    additionalIncome = None
    additionalDeduction = None
    withholdingEffectiveStartDate = None

    def __init__(self):
        self.code = None
        self.amount = None
        self.rate = None
        self.exemptions = None
        self.exemptionAmount = None
        self.exemptionPercent = None
        self.block = False
        self.type = None
        self.filingStatus = None
        self.headOfHousehold = False
        self.psdRate = None
        self.psdCode = None
        self.nonResidentAlienAdditionalFitwh = None
        self.applicableBirthYear = None
        self.adjustWithholding = None
        self.spouseWork = None
        self.livedInWorkedIn = None
        self.calculate = None
        self.numberOfDependents = None
        self.numberOfOtherDependents = None
        self.hasTwoIncomes = False
        self.additionalIncome = None
        self.additionalDeduction = None
        self.withholdingEffectiveStartDate = None


class EmployeeComplex:
    names = Names()
    ssn = None
    employeeId = None
    paycorNumber = None
    dept = None
    deptDescription = None
    payrollCode = None
    priorPaygroup = None
    status = None
    hireDate = None
    rehireDate = None
    birthDate = None
    sex = None
    termDate = None
    employeeType = None
    statusCode = None
    race = None
    maritalStatus = None
    country = None
    state = None
    zip = None
    city = None
    address1 = None
    address2 = None
    employmentStatusType = None
    pay = None
    accounts = []
    earningCodes = []
    taxCodes = []
    deductionCodes = []
    altCodes = []
    manager = Manager()
    jobTitle = None
    phoneMobile = None
    phoneHome = None
    phoneWork = None
    phoneWorkExtension = None
    emailHome = None
    emailWork = None
    unknownValues = None

    def __init__(self):
        self.names = Names()
        self.ssn = None
        self.employeeId = None
        self.paycorNumber = None
        self.dept = None
        self.deptDescription = None
        self.payrollCode = None
        self.priorPaygroup = None
        self.status = None
        self.hireDate = None
        self.rehireDate = None
        self.birthDate = None
        self.sex = None
        self.termDate = None
        self.employeeType = None
        self.statusCode = None
        self.race = None
        self.maritalStatus = None
        self.country = None
        self.state = None
        self.zip = None
        self.city = None
        self.address1 = None
        self.address2 = None
        self.employmentStatusType = None
        self.pay = Pay()
        self.accounts = []
        self.earningCodes = []
        self.taxCodes = []
        self.deductionCodes = []
        self.altCodes = []
        self.manager = Manager()
        self.jobTitle = None
        self.phoneMobile = None
        self.phoneHome = None
        self.phoneWork = None
        self.phoneWorkExtension = None
        self.emailHome = None
        self.emailWork = None
        self.unknownValues = {}


class PayRate:
    amount = None
    description = None
    type = None
    sequence = None
    effectiveDate = None
    reason = None

    def __init__(self):
        self.amount = None
        self.description = None
        self.type = None
        self.sequence = None
        self.effectiveDate = None
        self.reason = None


class EmployeeOrgImp:
    clientid = "",
    employeeNumber = "",
    person_firstName = "",
    person_middleName = "",
    person_lastName = "",
    person_prefix = "",
    person_suffix = "",
    departmentCode = "",
    payrollCode = "",
    paygroupDescription = "",
    primaryaddress_addressLine1 = "",
    primaryaddress_addressLine2 = "",
    primaryaddress_suite = "",
    primaryaddress_zip = "",
    primaryaddress_city = "",
    primaryaddress_state = "",
    primaryaddress_country = "",
    workPhone = "",
    workFaxNumber = "",
    workEmail = "",
    employmentStatus = "",
    hireDate = "",
    reHireDate = "",
    terminationDate = "",
    terminationReason = "",
    flsa = "",
    statusType = "",
    jobTitle = "",
    employeeType = "",
    person_birthDate = "",
    person_ssn = "",
    person_gender = "",
    person_homePhone = "",
    person_homeEmail = "",
    person_mobilePhone = "",
    person_ethnicity = "",
    person_maritalStatus = "",
    person_managerEmployeeNumber = "",
    person_managerClientId = "",
    annualHours = "",
    ownerOfficer = "",
    ownershipPercent = "",
    baseShift = "",
    keyEmployee = "",
    highlyCompensatedEmployee = "",
    familyMember = "",
    benefitClassification = "",
    person_maidenName = "",
    person_accredited = "",
    person_legalFirstName = "",
    person_legalLastName = "",
    person_veteranStatus = "",
    person_tobaccoUser = "",
    person_disabilityStatus = "",
    person_dischargeDate = "",
    person_isArmedForcesServiceMedalVeteran = "",
    person_isDisabledVeteran = "",
    person_isOtherProtectedVeteran = "",
    person_isRecentlySeperatedVeteran = "",
    person_isSpecialDisabledVeteran = "",
    person_isVietnamEra = ""

    def __init__(self):
        self.clientid = ""
        self.employeeNumber = ""
        self.person_firstName = ""
        self.person_middleName = ""
        self.person_lastName = ""
        self.person_prefix = ""
        self.person_suffix = ""
        self.departmentCode = ""
        self.payrollCode = ""
        self.paygroupDescription = ""
        self.primaryaddress_addressLine1 = ""
        self.primaryaddress_addressLine2 = ""
        self.primaryaddress_suite = ""
        self.primaryaddress_zip = ""
        self.primaryaddress_city = ""
        self.primaryaddress_state = ""
        self.primaryaddress_country = ""
        self.workPhone = ""
        self.workFaxNumber = ""
        self.workEmail = ""
        self.employmentStatus = ""
        self.hireDate = ""
        self.reHireDate = ""
        self.terminationDate = ""
        self.terminationReason = ""
        self.flsa = ""
        self.statusType = ""
        self.jobTitle = ""
        self.employeeType = ""
        self.person_birthDate = ""
        self.person_ssn = ""
        self.person_gender = ""
        self.person_homePhone = ""
        self.person_homeEmail = ""
        self.person_mobilePhone = ""
        self.person_ethnicity = ""
        self.person_maritalStatus = ""
        self.person_managerEmployeeNumber = ""
        self.person_managerClientId = ""
        self.annualHours = ""
        self.ownerOfficer = ""
        self.ownershipPercent = ""
        self.baseShift = ""
        self.keyEmployee = ""
        self.highlyCompensatedEmployee = ""
        self.familyMember = ""
        self.benefitClassification = ""
        self.person_maidenName = ""
        self.person_accredited = ""
        self.person_legalFirstName = ""
        self.person_legalLastName = ""
        self.person_veteranStatus = ""
        self.person_tobaccoUser = ""
        self.person_disabilityStatus = ""
        self.person_dischargeDate = ""
        self.person_isArmedForcesServiceMedalVeteran = ""
        self.person_isDisabledVeteran = ""
        self.person_isOtherProtectedVeteran = ""
        self.person_isRecentlySeperatedVeteran = ""
        self.person_isSpecialDisabledVeteran = ""
        self.person_isVietnamEra = ""


class EmployeePay:
    priorId = None
    firstName = None
    lastName = None
    checkDate = None
    checkNumber = None
    deptId = None
    itemType = None
    priorCode = None
    amount = None

    def __init__(self):
        self.priorId = None
        self.firstName = None
        self.lastName = None
        self.checkDate = None
        self.checkNumber = None
        self.deptId = None
        self.itemType = None
        self.priorCode = None
        self.amount = None

    def changeName(self, firstName, lastName):
        if firstName != None:
            self.firstName = firstName.replace(u"\u2212", "-")

        if lastName != None:
            self.lastName = lastName.replace(u"\u2212", "-")


class EmployeePriorPay:
    priorId = None
    checkDate = None
    checkNumber = None
    firstName = None
    lastName = None
    payItems = []
    fileId = None
    fileVersion = None
    fileName = None
    fileUploadDateTime = None
    fileUploadedByUser = None

    def __init__(self):
        self.priorId = None
        self.checkDate = None
        self.checkNumber = None
        self.firstName = None
        self.lastName = None
        self.payItems = []
        self.fileId = None
        self.fileVersion = None
        self.firstName = None
        self.fileUploadDateTime = None
        self.fileUploadedByUser = None

    def changeName(self, firstName, lastName):
        self.firstName = firstName
        self.lastName = lastName


class PayItem:
    itemType = None
    priorCode = None
    amount = None
    deptId = None

    def __init__(self):
        self.itemType = None
        self.priorCode = None
        self.amount = None
        self.deptId = None


def output_excel_multitab(filename, tabName, headers, employees):
    try:
        book = xlwt.Workbook()
        if os.path.isfile(filename):
            book1 = open_workbook(filename)
            for name in book1.sheet_names():
                if name != tabName:
                    target_sheet = book.add_sheet(name)
                    orig_sheet = book1.sheet_by_name(name)
                    numRow = orig_sheet.nrows
                    numCol = orig_sheet.ncols
                    for row in xrange(numRow):
                        rowList = orig_sheet.row_values(row)
                        for col in xrange(numCol):
                            oneVal = rowList[col]
                            target_sheet.write(row, col, oneVal)

        sh = book.add_sheet(tabName)

        header_mapping = {}

        for colIndex in range(0, len(headers)):
            sh.write(0, colIndex, headers[colIndex])
            header_mapping[headers[colIndex]] = colIndex

        row = 1
        for employee in employees:
            for key in employee.keys():
                sh.write(row, header_mapping[key], employee[key])
            row = row + 1

        book.save(filename)
    except Exception as E:
        t = 5

def split_city_state_zip(city_state_zip):
    city = city_state_zip
    state = ""
    zip_code = ""

    if "," not in city_state_zip:
        return city, state, zip_code
    first_comma = city_state_zip.index(",")
    city = city_state_zip[:first_comma]
    city_state_zip = city_state_zip[first_comma:]
    if " " not in city_state_zip:
        return city, state, zip_code
    last_space = city_state_zip.rfind(" ")
    zip_code = city_state_zip[last_space:].strip()
    state = city_state_zip[1:last_space].strip()
    return city, state, zip_code


def strip_if_not_none(value):
    if value is None:
        return value
    value = value.strip()
    if len(value) == 0:
        return None
    return value


def get_column_number(column_name, employee_header_columns):
    header_row = next((y for y in employee_header_columns if y[0].lower() == column_name.lower()), None)
    if header_row is None:
        return -1
    return header_row[1]


def usaddress_parse(full_address):
    import usaddress
    street_names = {'n', 's', 'e', 'w', 'N', 'S', 'E', 'W', 'ne', 'nw', 'se', 'sw',
                    'NE', 'NW', 'SE', 'SW', 'n.', 'N.', 'e.', 'E.', 'w.', 'W.', 's.', 'S.',
                    'north', 'south', 'east', 'west', 'northeast', 'northwest', 'southeast', 'southwest',
                    'SOUTH', 'NORTH', 'EAST', 'WEST', 'NORTHEAST', 'NORTHWEST', 'SOUTHEAST', 'SOUTHWEST',
                    'allee', 'alley', 'ally', 'aly', 'anex', 'annex', 'annx', 'anx', 'arc',
                    'arcade', 'av', 'ave', 'aven', 'avenu', 'avenue', 'avn', 'avnue', 'bayoo',
                    'bayou', 'bch', 'beach', 'bend', 'bg', 'bgs', 'blf', 'blfs', 'bluf',
                    'bluff', 'bluffs', 'blvd', 'bnd', 'bot', 'bottm', 'bottom', 'boul',
                    'boulevard', 'boulv', 'br', 'branch', 'brdge', 'brg', 'bridge', 'brk',
                    'brks', 'brnch', 'brook', 'brooks', 'btm', 'burg', 'burgs', 'byp', 'bypa',
                    'bypas', 'bypass', 'byps', 'byu', 'camp', 'canyn', 'canyon', 'cape',
                    'causeway', 'causwa', 'cen', 'cent', 'center', 'centers', 'centr',
                    'centre', 'cir', 'circ', 'circl', 'circle', 'circles', 'cirs', 'clb',
                    'clf', 'clfs', 'cliff', 'cliffs', 'club', 'cmn', 'cmns', 'cmp', 'cnter',
                    'cntr', 'cnyn', 'common', 'commons', 'cor', 'corner', 'corners', 'cors',
                    'course', 'court', 'courts', 'cove', 'coves', 'cp', 'cpe', 'crcl', 'crcle',
                    'creek', 'cres', 'crescent', 'crest', 'crk', 'crossing', 'crossroad',
                    'crossroads', 'crse', 'crsent', 'crsnt', 'crssng', 'crst', 'cswy', 'ct',
                    'ctr', 'ctrs', 'cts', 'curv', 'curve', 'cv', 'cvs', 'cyn', 'dale', 'dam',
                    'div', 'divide', 'dl', 'dm', 'dr', 'driv', 'drive', 'drives', 'drs', 'drv',
                    'dv', 'dvd', 'est', 'estate', 'estates', 'ests', 'exp', 'expr', 'express',
                    'expressway', 'expw', 'expy', 'ext', 'extension', 'extensions', 'extn',
                    'extnsn', 'exts', 'fall', 'falls', 'ferry', 'field', 'fields', 'flat',
                    'flats', 'fld', 'flds', 'fls', 'flt', 'flts', 'ford', 'fords', 'forest',
                    'forests', 'forg', 'forge', 'forges', 'fork', 'forks', 'fort', 'frd',
                    'frds', 'freeway', 'freewy', 'frg', 'frgs', 'frk', 'frks', 'frry', 'frst',
                    'frt', 'frway', 'frwy', 'fry', 'ft', 'fwy', 'garden', 'gardens', 'gardn',
                    'gateway', 'gatewy', 'gatway', 'gdn', 'gdns', 'glen', 'glens', 'gln',
                    'glns', 'grden', 'grdn', 'grdns', 'green', 'greens', 'grn', 'grns', 'grov',
                    'grove', 'groves', 'grv', 'grvs', 'gtway', 'gtwy', 'harb', 'harbor',
                    'harbors', 'harbr', 'haven', 'hbr', 'hbrs', 'heights', 'highway', 'highwy',
                    'hill', 'hills', 'hiway', 'hiwy', 'hl', 'hllw', 'hls', 'hollow', 'hollows',
                    'holw', 'holws', 'hrbor', 'ht', 'hts', 'hvn', 'hway', 'hwy', 'inlet',
                    'inlt', 'is', 'island', 'islands', 'isle', 'isles', 'islnd', 'islnds',
                    'iss', 'jct', 'jction', 'jctn', 'jctns', 'jcts', 'junction', 'junctions',
                    'junctn', 'juncton', 'key', 'keys', 'knl', 'knls', 'knol', 'knoll',
                    'knolls', 'ky', 'kys', 'lake', 'lakes', 'land', 'landing', 'lane', 'lck',
                    'lcks', 'ldg', 'ldge', 'lf', 'lgt', 'lgts', 'light', 'lights', 'lk', 'lks',
                    'ln', 'lndg', 'lndng', 'loaf', 'lock', 'locks', 'lodg', 'lodge', 'loop',
                    'loops', 'mall', 'manor', 'manors', 'mdw', 'mdws', 'meadow', 'meadows',
                    'medows', 'mews', 'mill', 'mills', 'mission', 'missn', 'ml', 'mls', 'mnr',
                    'mnrs', 'mnt', 'mntain', 'mntn', 'mntns', 'motorway', 'mount', 'mountain',
                    'mountains', 'mountin', 'msn', 'mssn', 'mt', 'mtin', 'mtn', 'mtns', 'mtwy',
                    'nck', 'neck', 'opas', 'orch', 'orchard', 'orchrd', 'oval', 'overpass',
                    'ovl', 'park', 'parks', 'parkway', 'parkways', 'parkwy', 'pass', 'passage',
                    'path', 'paths', 'pike', 'pikes', 'pine', 'pines', 'pkway', 'pkwy',
                    'pkwys', 'pky', 'pl', 'place', 'plain', 'plains', 'plaza', 'pln', 'plns',
                    'plz', 'plza', 'pne', 'pnes', 'point', 'points', 'port', 'ports', 'pr',
                    'prairie', 'prk', 'prr', 'prt', 'prts', 'psge', 'pt', 'pts', 'rad',
                    'radial', 'radiel', 'radl', 'ramp', 'ranch', 'ranches', 'rapid', 'rapids',
                    'rd', 'rdg', 'rdge', 'rdgs', 'rds', 'rest', 'ridge', 'ridges', 'riv',
                    'river', 'rivr', 'rnch', 'rnchs', 'road', 'roads', 'route', 'row', 'rpd',
                    'rpds', 'rst', 'rte', 'rue', 'run', 'rvr', 'shl', 'shls', 'shoal',
                    'shoals', 'shoar', 'shoars', 'shore', 'shores', 'shr', 'shrs', 'skwy',
                    'skyway', 'smt', 'spg', 'spgs', 'spng', 'spngs', 'spring', 'springs',
                    'sprng', 'sprngs', 'spur', 'spurs', 'sq', 'sqr', 'sqre', 'sqrs', 'sqs',
                    'squ', 'square', 'squares', 'st', 'sta', 'station', 'statn', 'stn', 'str',
                    'stra', 'strav', 'straven', 'stravenue', 'stravn', 'stream', 'street',
                    'streets', 'streme', 'strm', 'strt', 'strvn', 'strvnue', 'sts', 'sumit',
                    'sumitt', 'summit', 'ter', 'terr', 'terrace', 'throughway', 'tpke',
                    'trace', 'traces', 'track', 'tracks', 'trafficway', 'trail', 'trailer',
                    'trails', 'trak', 'trce', 'trfy', 'trk', 'trks', 'trl', 'trlr', 'trlrs',
                    'trls', 'trnpk', 'trwy', 'tunel', 'tunl', 'tunls', 'tunnel', 'tunnels',
                    'tunnl', 'turnpike', 'turnpk', 'un', 'underpass', 'union', 'unions', 'uns',
                    'upas', 'valley', 'valleys', 'vally', 'vdct', 'via', 'viadct', 'viaduct',
                    'view', 'views', 'vill', 'villag', 'village', 'villages', 'ville', 'villg',
                    'villiage', 'vis', 'vist', 'vista', 'vl', 'vlg', 'vlgs', 'vlly', 'vly',
                    'vlys', 'vst', 'vsta', 'vw', 'vws', 'walk', 'walks', 'wall', 'way', 'ways',
                    'well', 'wells', 'wl', 'wls', 'wy', 'xing', 'xrd', 'xrds',
                    'ALLEE', 'ALLEY', 'ALLY', 'ALY', 'ANEX', 'ANNEX', 'ANNX', 'ANX', 'ARC',
                    'ARCADE', 'AV', 'AVE', 'AVEN', 'AVENU', 'AVENUE', 'AVN', 'AVNUE', 'BAYOO',
                    'BAYOU', 'BCH', 'BEACH', 'BEND', 'BG', 'BGS', 'BLF', 'BLFS', 'BLUF',
                    'BLUFF', 'BLUFFS', 'BLVD', 'BND', 'BOT', 'BOTTM', 'BOTTOM', 'BOUL',
                    'BOULEVARD', 'BOULV', 'BR', 'BRANCH', 'BRDGE', 'BRG', 'BRIDGE', 'BRK',
                    'BRKS', 'BRNCH', 'BROOK', 'BROOKS', 'BTM', 'BURG', 'BURGS', 'BYP', 'BYPA',
                    'BYPAS', 'BYPASS', 'BYPS', 'BYU', 'CAMP', 'CANYN', 'CANYON', 'CAPE',
                    'CAUSEWAY', 'CAUSWA', 'CEN', 'CENT', 'CENTER', 'CENTERS', 'CENTR',
                    'CENTRE', 'CIR', 'CIRC', 'CIRCL', 'CIRCLE', 'CIRCLES', 'CIRS', 'CLB',
                    'CLF', 'CLFS', 'CLIFF', 'CLIFFS', 'CLUB', 'CMN', 'CMNS', 'CMP', 'CNTER',
                    'CNTR', 'CNYN', 'COMMON', 'COMMONS', 'COR', 'CORNER', 'CORNERS', 'CORS',
                    'COURSE', 'COURT', 'COURTS', 'COVE', 'COVES', 'CP', 'CPE', 'CRCL', 'CRCLE',
                    'CREEK', 'CRES', 'CRESCENT', 'CREST', 'CRK', 'CROSSING', 'CROSSROAD',
                    'CROSSROADS', 'CRSE', 'CRSENT', 'CRSNT', 'CRSSNG', 'CRST', 'CSWY', 'CT',
                    'CTR', 'CTRS', 'CTS', 'CURV', 'CURVE', 'CV', 'CVS', 'CYN', 'DALE', 'DAM',
                    'DIV', 'DIVIDE', 'DL', 'DM', 'DR', 'DRIV', 'DRIVE', 'DRIVES', 'DRS', 'DRV',
                    'DV', 'DVD', 'EST', 'ESTATE', 'ESTATES', 'ESTS', 'EXP', 'EXPR', 'EXPRESS',
                    'EXPRESSWAY', 'EXPW', 'EXPY', 'EXT', 'EXTENSION', 'EXTENSIONS', 'EXTN',
                    'EXTNSN', 'EXTS', 'FALL', 'FALLS', 'FERRY', 'FIELD', 'FIELDS', 'FLAT',
                    'FLATS', 'FLD', 'FLDS', 'FLS', 'FLT', 'FLTS', 'FORD', 'FORDS', 'FOREST',
                    'FORESTS', 'FORG', 'FORGE', 'FORGES', 'FORK', 'FORKS', 'FORT', 'FRD',
                    'FRDS', 'FREEWAY', 'FREEWY', 'FRG', 'FRGS', 'FRK', 'FRKS', 'FRRY', 'FRST',
                    'FRT', 'FRWAY', 'FRWY', 'FRY', 'FT', 'FWY', 'GARDEN', 'GARDENS', 'GARDN',
                    'GATEWAY', 'GATEWY', 'GATWAY', 'GDN', 'GDNS', 'GLEN', 'GLENS', 'GLN',
                    'GLNS', 'GRDEN', 'GRDN', 'GRDNS', 'GREEN', 'GREENS', 'GRN', 'GRNS', 'GROV',
                    'GROVE', 'GROVES', 'GRV', 'GRVS', 'GTWAY', 'GTWY', 'HARB', 'HARBOR',
                    'HARBORS', 'HARBR', 'HAVEN', 'HBR', 'HBRS', 'HEIGHTS', 'HIGHWAY', 'HIGHWY',
                    'HILL', 'HILLS', 'HIWAY', 'HIWY', 'HL', 'HLLW', 'HLS', 'HOLLOW', 'HOLLOWS',
                    'HOLW', 'HOLWS', 'HRBOR', 'HT', 'HTS', 'HVN', 'HWAY', 'HWY', 'INLET',
                    'INLT', 'IS', 'ISLAND', 'ISLANDS', 'ISLE', 'ISLES', 'ISLND', 'ISLNDS',
                    'ISS', 'JCT', 'JCTION', 'JCTN', 'JCTNS', 'JCTS', 'JUNCTION', 'JUNCTIONS',
                    'JUNCTN', 'JUNCTON', 'KEY', 'KEYS', 'KNL', 'KNLS', 'KNOL', 'KNOLL',
                    'KNOLLS', 'KY', 'KYS', 'LAKE', 'LAKES', 'LAND', 'LANDING', 'LANE', 'LCK',
                    'LCKS', 'LDG', 'LDGE', 'LF', 'LGT', 'LGTS', 'LIGHT', 'LIGHTS', 'LK', 'LKS',
                    'LN', 'LNDG', 'LNDNG', 'LOAF', 'LOCK', 'LOCKS', 'LODG', 'LODGE', 'LOOP',
                    'LOOPS', 'MALL', 'MANOR', 'MANORS', 'MDW', 'MDWS', 'MEADOW', 'MEADOWS',
                    'MEDOWS', 'MEWS', 'MILL', 'MILLS', 'MISSION', 'MISSN', 'ML', 'MLS', 'MNR',
                    'MNRS', 'MNT', 'MNTAIN', 'MNTN', 'MNTNS', 'MOTORWAY', 'MOUNT', 'MOUNTAIN',
                    'MOUNTAINS', 'MOUNTIN', 'MSN', 'MSSN', 'MT', 'MTIN', 'MTN', 'MTNS', 'MTWY',
                    'NCK', 'NECK', 'OPAS', 'ORCH', 'ORCHARD', 'ORCHRD', 'OVAL', 'OVERPASS',
                    'OVL', 'PARK', 'PARKS', 'PARKWAY', 'PARKWAYS', 'PARKWY', 'PASS', 'PASSAGE',
                    'PATH', 'PATHS', 'PIKE', 'PIKES', 'PINE', 'PINES', 'PKWAY', 'PKWY',
                    'PKWYS', 'PKY', 'PL', 'PLACE', 'PLAIN', 'PLAINS', 'PLAZA', 'PLN', 'PLNS',
                    'PLZ', 'PLZA', 'PNE', 'PNES', 'POINT', 'POINTS', 'PORT', 'PORTS', 'PR',
                    'PRAIRIE', 'PRK', 'PRR', 'PRT', 'PRTS', 'PSGE', 'PT', 'PTS', 'RAD',
                    'RADIAL', 'RADIEL', 'RADL', 'RAMP', 'RANCH', 'RANCHES', 'RAPID', 'RAPIDS',
                    'RD', 'RDG', 'RDGE', 'RDGS', 'RDS', 'REST', 'RIDGE', 'RIDGES', 'RIV',
                    'RIVER', 'RIVR', 'RNCH', 'RNCHS', 'ROAD', 'ROADS', 'ROUTE', 'ROW', 'RPD',
                    'RPDS', 'RST', 'RTE', 'RUE', 'RUN', 'RVR', 'SHL', 'SHLS', 'SHOAL',
                    'SHOALS', 'SHOAR', 'SHOARS', 'SHORE', 'SHORES', 'SHR', 'SHRS', 'SKWY',
                    'SKYWAY', 'SMT', 'SPG', 'SPGS', 'SPNG', 'SPNGS', 'SPRING', 'SPRINGS',
                    'SPRNG', 'SPRNGS', 'SPUR', 'SPURS', 'SQ', 'SQR', 'SQRE', 'SQRS', 'SQS',
                    'SQU', 'SQUARE', 'SQUARES', 'ST', 'STA', 'STATION', 'STATN', 'STN', 'STR',
                    'STRA', 'STRAV', 'STRAVEN', 'STRAVENUE', 'STRAVN', 'STREAM', 'STREET',
                    'STREETS', 'STREME', 'STRM', 'STRT', 'STRVN', 'STRVNUE', 'STS', 'SUMIT',
                    'SUMITT', 'SUMMIT', 'TER', 'TERR', 'TERRACE', 'THROUGHWAY', 'TPKE',
                    'TRACE', 'TRACES', 'TRACK', 'TRACKS', 'TRAFFICWAY', 'TRAIL', 'TRAILER',
                    'TRAILS', 'TRAK', 'TRCE', 'TRFY', 'TRK', 'TRKS', 'TRL', 'TRLR', 'TRLRS',
                    'TRLS', 'TRNPK', 'TRWY', 'TUNEL', 'TUNL', 'TUNLS', 'TUNNEL', 'TUNNELS',
                    'TUNNL', 'TURNPIKE', 'TURNPK', 'UN', 'UNDERPASS', 'UNION', 'UNIONS', 'UNS',
                    'UPAS', 'VALLEY', 'VALLEYS', 'VALLY', 'VDCT', 'VIA', 'VIADCT', 'VIADUCT',
                    'VIEW', 'VIEWS', 'VILL', 'VILLAG', 'VILLAGE', 'VILLAGES', 'VILLE', 'VILLG',
                    'VILLIAGE', 'VIS', 'VIST', 'VISTA', 'VL', 'VLG', 'VLGS', 'VLLY', 'VLY',
                    'VLYS', 'VST', 'VSTA', 'VW', 'VWS', 'WALK', 'WALKS', 'WALL', 'WAY', 'WAYS',
                    'WELL', 'WELLS', 'WL', 'WLS', 'WY', 'XING', 'XRD', 'XRDS',
                    'allee.', 'alley.', 'ally.', 'aly.', 'anex.', 'annex.', 'annx.', 'anx.', 'arc',
                    'arcade.', 'av.', 'ave.', 'aven.', 'avenu.', 'avenue.', 'avn.', 'avnue.', 'bayoo',
                    'bayou.', 'bch.', 'beach.', 'bend.', 'bg.', 'bgs.', 'blf.', 'blfs.', 'bluf',
                    'bluff.', 'bluffs.', 'blvd.', 'bnd.', 'bot.', 'bottm.', 'bottom.', 'boul',
                    'boulevard.', 'boulv.', 'br.', 'branch.', 'brdge.', 'brg.', 'bridge.', 'brk',
                    'brks.', 'brnch.', 'brook.', 'brooks.', 'btm.', 'burg.', 'burgs.', 'byp.', 'bypa',
                    'bypas.', 'bypass.', 'byps.', 'byu.', 'camp.', 'canyn.', 'canyon.', 'cape',
                    'causeway.', 'causwa.', 'cen.', 'cent.', 'center.', 'centers.', 'centr',
                    'centre.', 'cir.', 'circ.', 'circl.', 'circle.', 'circles.', 'cirs.', 'clb',
                    'clf.', 'clfs.', 'cliff.', 'cliffs.', 'club.', 'cmn.', 'cmns.', 'cmp.', 'cnter',
                    'cntr.', 'cnyn.', 'common.', 'commons.', 'cor.', 'corner.', 'corners.', 'cors',
                    'course.', 'court.', 'courts.', 'cove.', 'coves.', 'cp.', 'cpe.', 'crcl.', 'crcle',
                    'creek.', 'cres.', 'crescent.', 'crest.', 'crk.', 'crossing.', 'crossroad',
                    'crossroads.', 'crse.', 'crsent.', 'crsnt.', 'crssng.', 'crst.', 'cswy.', 'ct',
                    'ctr.', 'ctrs.', 'cts.', 'curv.', 'curve.', 'cv.', 'cvs.', 'cyn.', 'dale.', 'dam',
                    'div.', 'divide.', 'dl.', 'dm.', 'dr.', 'driv.', 'drive.', 'drives.', 'drs.', 'drv',
                    'dv.', 'dvd.', 'est.', 'estate.', 'estates.', 'ests.', 'exp.', 'expr.', 'express',
                    'expressway.', 'expw.', 'expy.', 'ext.', 'extension.', 'extensions.', 'extn',
                    'extnsn.', 'exts.', 'fall.', 'falls.', 'ferry.', 'field.', 'fields.', 'flat',
                    'flats.', 'fld.', 'flds.', 'fls.', 'flt.', 'flts.', 'ford.', 'fords.', 'forest',
                    'forests.', 'forg.', 'forge.', 'forges.', 'fork.', 'forks.', 'fort.', 'frd',
                    'frds.', 'freeway.', 'freewy.', 'frg.', 'frgs.', 'frk.', 'frks.', 'frry.', 'frst',
                    'frt.', 'frway.', 'frwy.', 'fry.', 'ft.', 'fwy.', 'garden.', 'gardens.', 'gardn',
                    'gateway.', 'gatewy.', 'gatway.', 'gdn.', 'gdns.', 'glen.', 'glens.', 'gln',
                    'glns.', 'grden.', 'grdn.', 'grdns.', 'green.', 'greens.', 'grn.', 'grns.', 'grov',
                    'grove.', 'groves.', 'grv.', 'grvs.', 'gtway.', 'gtwy.', 'harb.', 'harbor',
                    'harbors.', 'harbr.', 'haven.', 'hbr.', 'hbrs.', 'heights.', 'highway.', 'highwy',
                    'hill.', 'hills.', 'hiway.', 'hiwy.', 'hl.', 'hllw.', 'hls.', 'hollow.', 'hollows',
                    'holw.', 'holws.', 'hrbor.', 'ht.', 'hts.', 'hvn.', 'hway.', 'hwy.', 'inlet',
                    'inlt.', 'is.', 'island.', 'islands.', 'isle.', 'isles.', 'islnd.', 'islnds',
                    'iss.', 'jct.', 'jction.', 'jctn.', 'jctns.', 'jcts.', 'junction.', 'junctions',
                    'junctn.', 'juncton.', 'key.', 'keys.', 'knl.', 'knls.', 'knol.', 'knoll',
                    'knolls.', 'ky.', 'kys.', 'lake.', 'lakes.', 'land.', 'landing.', 'lane.', 'lck',
                    'lcks.', 'ldg.', 'ldge.', 'lf.', 'lgt.', 'lgts.', 'light.', 'lights.', 'lk.', 'lks',
                    'ln.', 'lndg.', 'lndng.', 'loaf.', 'lock.', 'locks.', 'lodg.', 'lodge.', 'loop',
                    'loops.', 'mall.', 'manor.', 'manors.', 'mdw.', 'mdws.', 'meadow.', 'meadows',
                    'medows.', 'mews.', 'mill.', 'mills.', 'mission.', 'missn.', 'ml.', 'mls.', 'mnr',
                    'mnrs.', 'mnt.', 'mntain.', 'mntn.', 'mntns.', 'motorway.', 'mount.', 'mountain',
                    'mountains.', 'mountin.', 'msn.', 'mssn.', 'mt.', 'mtin.', 'mtn.', 'mtns.', 'mtwy',
                    'nck.', 'neck.', 'opas.', 'orch.', 'orchard.', 'orchrd.', 'oval.', 'overpass',
                    'ovl.', 'park.', 'parks.', 'parkway.', 'parkways.', 'parkwy.', 'pass.', 'passage',
                    'path.', 'paths.', 'pike.', 'pikes.', 'pine.', 'pines.', 'pkway.', 'pkwy',
                    'pkwys.', 'pky.', 'pl.', 'place.', 'plain.', 'plains.', 'plaza.', 'pln.', 'plns',
                    'plz.', 'plza.', 'pne.', 'pnes.', 'point.', 'points.', 'port.', 'ports.', 'pr',
                    'prairie.', 'prk.', 'prr.', 'prt.', 'prts.', 'psge.', 'pt.', 'pts.', 'rad',
                    'radial.', 'radiel.', 'radl.', 'ramp.', 'ranch.', 'ranches.', 'rapid.', 'rapids',
                    'rd.', 'rdg.', 'rdge.', 'rdgs.', 'rds.', 'rest.', 'ridge.', 'ridges.', 'riv',
                    'river.', 'rivr.', 'rnch.', 'rnchs.', 'road.', 'roads.', 'route.', 'row.', 'rpd',
                    'rpds.', 'rst.', 'rte.', 'rue.', 'run.', 'rvr.', 'shl.', 'shls.', 'shoal',
                    'shoals.', 'shoar.', 'shoars.', 'shore.', 'shores.', 'shr.', 'shrs.', 'skwy',
                    'skyway.', 'smt.', 'spg.', 'spgs.', 'spng.', 'spngs.', 'spring.', 'springs',
                    'sprng.', 'sprngs.', 'spur.', 'spurs.', 'sq.', 'sqr.', 'sqre.', 'sqrs.', 'sqs',
                    'squ.', 'square.', 'squares.', 'st.', 'sta.', 'station.', 'statn.', 'stn.', 'str',
                    'stra.', 'strav.', 'straven.', 'stravenue.', 'stravn.', 'stream.', 'street',
                    'streets.', 'streme.', 'strm.', 'strt.', 'strvn.', 'strvnue.', 'sts.', 'sumit',
                    'sumitt.', 'summit.', 'ter.', 'terr.', 'terrace.', 'throughway.', 'tpke',
                    'trace.', 'traces.', 'track.', 'tracks.', 'trafficway.', 'trail.', 'trailer',
                    'trails.', 'trak.', 'trce.', 'trfy.', 'trk.', 'trks.', 'trl.', 'trlr.', 'trlrs',
                    'trls.', 'trnpk.', 'trwy.', 'tunel.', 'tunl.', 'tunls.', 'tunnel.', 'tunnels',
                    'tunnl.', 'turnpike.', 'turnpk.', 'un.', 'underpass.', 'union.', 'unions.', 'uns',
                    'upas.', 'valley.', 'valleys.', 'vally.', 'vdct.', 'via.', 'viadct.', 'viaduct',
                    'view.', 'views.', 'vill.', 'villag.', 'village.', 'villages.', 'ville.', 'villg',
                    'villiage.', 'vis.', 'vist.', 'vista.', 'vl.', 'vlg.', 'vlgs.', 'vlly.', 'vly',
                    'vlys.', 'vst.', 'vsta.', 'vw.', 'vws.', 'walk.', 'walks.', 'wall.', 'way.', 'ways',
                    'well.', 'wells.', 'wl.', 'wls.', 'wy.', 'xing.', 'xrd.', 'xrds',
                    'ALLEE.', 'ALLEY.', 'ALLY.', 'ALY.', 'ANEX.', 'ANNEX.', 'ANNX.', 'ANX.', 'ARC',
                    'ARCADE.', 'AV.', 'AVE.', 'AVEN.', 'AVENU.', 'AVENUE.', 'AVN.', 'AVNUE.', 'BAYOO',
                    'BAYOU.', 'BCH.', 'BEACH.', 'BEND.', 'BG.', 'BGS.', 'BLF.', 'BLFS.', 'BLUF',
                    'BLUFF.', 'BLUFFS.', 'BLVD.', 'BND.', 'BOT.', 'BOTTM.', 'BOTTOM.', 'BOUL',
                    'BOULEVARD.', 'BOULV.', 'BR.', 'BRANCH.', 'BRDGE.', 'BRG.', 'BRIDGE.', 'BRK',
                    'BRKS.', 'BRNCH.', 'BROOK.', 'BROOKS.', 'BTM.', 'BURG.', 'BURGS.', 'BYP.', 'BYPA',
                    'BYPAS.', 'BYPASS.', 'BYPS.', 'BYU.', 'CAMP.', 'CANYN.', 'CANYON.', 'CAPE',
                    'CAUSEWAY.', 'CAUSWA.', 'CEN.', 'CENT.', 'CENTER.', 'CENTERS.', 'CENTR',
                    'CENTRE.', 'CIR.', 'CIRC.', 'CIRCL.', 'CIRCLE.', 'CIRCLES.', 'CIRS.', 'CLB',
                    'CLF.', 'CLFS.', 'CLIFF.', 'CLIFFS.', 'CLUB.', 'CMN.', 'CMNS.', 'CMP.', 'CNTER',
                    'CNTR.', 'CNYN.', 'COMMON.', 'COMMONS.', 'COR.', 'CORNER.', 'CORNERS.', 'CORS',
                    'COURSE.', 'COURT.', 'COURTS.', 'COVE.', 'COVES.', 'CP.', 'CPE.', 'CRCL.', 'CRCLE',
                    'CREEK.', 'CRES.', 'CRESCENT.', 'CREST.', 'CRK.', 'CROSSING.', 'CROSSROAD',
                    'CROSSROADS.', 'CRSE.', 'CRSENT.', 'CRSNT.', 'CRSSNG.', 'CRST.', 'CSWY.', 'CT',
                    'CTR.', 'CTRS.', 'CTS.', 'CURV.', 'CURVE.', 'CV.', 'CVS.', 'CYN.', 'DALE.', 'DAM',
                    'DIV.', 'DIVIDE.', 'DL.', 'DM.', 'DR.', 'DRIV.', 'DRIVE.', 'DRIVES.', 'DRS.', 'DRV',
                    'DV.', 'DVD.', 'EST.', 'ESTATE.', 'ESTATES.', 'ESTS.', 'EXP.', 'EXPR.', 'EXPRESS',
                    'EXPRESSWAY.', 'EXPW.', 'EXPY.', 'EXT.', 'EXTENSION.', 'EXTENSIONS.', 'EXTN',
                    'EXTNSN.', 'EXTS.', 'FALL.', 'FALLS.', 'FERRY.', 'FIELD.', 'FIELDS.', 'FLAT',
                    'FLATS.', 'FLD.', 'FLDS.', 'FLS.', 'FLT.', 'FLTS.', 'FORD.', 'FORDS.', 'FOREST',
                    'FORESTS.', 'FORG.', 'FORGE.', 'FORGES.', 'FORK.', 'FORKS.', 'FORT.', 'FRD',
                    'FRDS.', 'FREEWAY.', 'FREEWY.', 'FRG.', 'FRGS.', 'FRK.', 'FRKS.', 'FRRY.', 'FRST',
                    'FRT.', 'FRWAY.', 'FRWY.', 'FRY.', 'FT.', 'FWY.', 'GARDEN.', 'GARDENS.', 'GARDN',
                    'GATEWAY.', 'GATEWY.', 'GATWAY.', 'GDN.', 'GDNS.', 'GLEN.', 'GLENS.', 'GLN',
                    'GLNS.', 'GRDEN.', 'GRDN.', 'GRDNS.', 'GREEN.', 'GREENS.', 'GRN.', 'GRNS.', 'GROV',
                    'GROVE.', 'GROVES.', 'GRV.', 'GRVS.', 'GTWAY.', 'GTWY.', 'HARB.', 'HARBOR',
                    'HARBORS.', 'HARBR.', 'HAVEN.', 'HBR.', 'HBRS.', 'HEIGHTS.', 'HIGHWAY.', 'HIGHWY',
                    'HILL.', 'HILLS.', 'HIWAY.', 'HIWY.', 'HL.', 'HLLW.', 'HLS.', 'HOLLOW.', 'HOLLOWS',
                    'HOLW.', 'HOLWS.', 'HRBOR.', 'HT.', 'HTS.', 'HVN.', 'HWAY.', 'HWY.', 'INLET',
                    'INLT.', 'IS.', 'ISLAND.', 'ISLANDS.', 'ISLE.', 'ISLES.', 'ISLND.', 'ISLNDS',
                    'ISS.', 'JCT.', 'JCTION.', 'JCTN.', 'JCTNS.', 'JCTS.', 'JUNCTION.', 'JUNCTIONS',
                    'JUNCTN.', 'JUNCTON.', 'KEY.', 'KEYS.', 'KNL.', 'KNLS.', 'KNOL.', 'KNOLL',
                    'KNOLLS.', 'KY.', 'KYS.', 'LAKE.', 'LAKES.', 'LAND.', 'LANDING.', 'LANE.', 'LCK',
                    'LCKS.', 'LDG.', 'LDGE.', 'LF.', 'LGT.', 'LGTS.', 'LIGHT.', 'LIGHTS.', 'LK.', 'LKS',
                    'LN.', 'LNDG.', 'LNDNG.', 'LOAF.', 'LOCK.', 'LOCKS.', 'LODG.', 'LODGE.', 'LOOP',
                    'LOOPS.', 'MALL.', 'MANOR.', 'MANORS.', 'MDW.', 'MDWS.', 'MEADOW.', 'MEADOWS',
                    'MEDOWS.', 'MEWS.', 'MILL.', 'MILLS.', 'MISSION.', 'MISSN.', 'ML.', 'MLS.', 'MNR',
                    'MNRS.', 'MNT.', 'MNTAIN.', 'MNTN.', 'MNTNS.', 'MOTORWAY.', 'MOUNT.', 'MOUNTAIN',
                    'MOUNTAINS.', 'MOUNTIN.', 'MSN.', 'MSSN.', 'MT.', 'MTIN.', 'MTN.', 'MTNS.', 'MTWY',
                    'NCK.', 'NECK.', 'OPAS.', 'ORCH.', 'ORCHARD.', 'ORCHRD.', 'OVAL.', 'OVERPASS',
                    'OVL.', 'PARK.', 'PARKS.', 'PARKWAY.', 'PARKWAYS.', 'PARKWY.', 'PASS.', 'PASSAGE',
                    'PATH.', 'PATHS.', 'PIKE.', 'PIKES.', 'PINE.', 'PINES.', 'PKWAY.', 'PKWY',
                    'PKWYS.', 'PKY.', 'PL.', 'PLACE.', 'PLAIN.', 'PLAINS.', 'PLAZA.', 'PLN.', 'PLNS',
                    'PLZ.', 'PLZA.', 'PNE.', 'PNES.', 'POINT.', 'POINTS.', 'PORT.', 'PORTS.', 'PR',
                    'PRAIRIE.', 'PRK.', 'PRR.', 'PRT.', 'PRTS.', 'PSGE.', 'PT.', 'PTS.', 'RAD',
                    'RADIAL.', 'RADIEL.', 'RADL.', 'RAMP.', 'RANCH.', 'RANCHES.', 'RAPID.', 'RAPIDS',
                    'RD.', 'RDG.', 'RDGE.', 'RDGS.', 'RDS.', 'REST.', 'RIDGE.', 'RIDGES.', 'RIV',
                    'RIVER.', 'RIVR.', 'RNCH.', 'RNCHS.', 'ROAD.', 'ROADS.', 'ROUTE.', 'ROW.', 'RPD',
                    'RPDS.', 'RST.', 'RTE.', 'RUE.', 'RUN.', 'RVR.', 'SHL.', 'SHLS.', 'SHOAL',
                    'SHOALS.', 'SHOAR.', 'SHOARS.', 'SHORE.', 'SHORES.', 'SHR.', 'SHRS.', 'SKWY',
                    'SKYWAY.', 'SMT.', 'SPG.', 'SPGS.', 'SPNG.', 'SPNGS.', 'SPRING.', 'SPRINGS',
                    'SPRNG.', 'SPRNGS.', 'SPUR.', 'SPURS.', 'SQ.', 'SQR.', 'SQRE.', 'SQRS.', 'SQS',
                    'SQU.', 'SQUARE.', 'SQUARES.', 'ST.', 'STA.', 'STATION.', 'STATN.', 'STN.', 'STR',
                    'STRA.', 'STRAV.', 'STRAVEN.', 'STRAVENUE.', 'STRAVN.', 'STREAM.', 'STREET',
                    'STREETS.', 'STREME.', 'STRM.', 'STRT.', 'STRVN.', 'STRVNUE.', 'STS.', 'SUMIT',
                    'SUMITT.', 'SUMMIT.', 'TER.', 'TERR.', 'TERRACE.', 'THROUGHWAY.', 'TPKE',
                    'TRACE.', 'TRACES.', 'TRACK.', 'TRACKS.', 'TRAFFICWAY.', 'TRAIL.', 'TRAILER',
                    'TRAILS.', 'TRAK.', 'TRCE.', 'TRFY.', 'TRK.', 'TRKS.', 'TRL.', 'TRLR.', 'TRLRS',
                    'TRLS.', 'TRNPK.', 'TRWY.', 'TUNEL.', 'TUNL.', 'TUNLS.', 'TUNNEL.', 'TUNNELS',
                    'TUNNL.', 'TURNPIKE.', 'TURNPK.', 'UN.', 'UNDERPASS.', 'UNION.', 'UNIONS.', 'UNS',
                    'UPAS.', 'VALLEY.', 'VALLEYS.', 'VALLY.', 'VDCT.', 'VIA.', 'VIADCT.', 'VIADUCT',
                    'VIEW.', 'VIEWS.', 'VILL.', 'VILLAG.', 'VILLAGE.', 'VILLAGES.', 'VILLE.', 'VILLG',
                    'VILLIAGE.', 'VIS.', 'VIST.', 'VISTA.', 'VL.', 'VLG.', 'VLGS.', 'VLLY.', 'VLY',
                    'VLYS.', 'VST.', 'VSTA.', 'VW.', 'VWS.', 'WALK.', 'WALKS.', 'WALL.', 'WAY.', 'WAYS',
                    'WELL.', 'WELLS.', 'WL.', 'WLS.', 'WY.', 'XING.', 'XRD.', 'XRDS',
                    'Allee.', 'Alley.', 'Ally.', 'Aly.', 'Anex.', 'Annex.', 'Annx.', 'Anx.', 'Arc',
                    'Arcade.', 'Av.', 'Ave.', 'Aven.', 'Avenu.', 'Avenue.', 'Avn.', 'Avnue.', 'Bayoo',
                    'Bayou.', 'Bch.', 'Beach.', 'Bend.', 'Bg.', 'Bgs.', 'Blf.', 'Blfs.', 'Bluf',
                    'Bluff.', 'Bluffs.', 'Blvd.', 'Bnd.', 'Bot.', 'Bottm.', 'Bottom.', 'Boul',
                    'Boulevard.', 'Boulv.', 'Br.', 'Branch.', 'Brdge.', 'Brg.', 'Bridge.', 'Brk',
                    'Brks.', 'Brnch.', 'Brook.', 'Brooks.', 'Btm.', 'Burg.', 'Burgs.', 'Byp.', 'Bypa',
                    'Bypas.', 'Bypass.', 'Byps.', 'Byu.', 'Camp.', 'Canyn.', 'Canyon.', 'Cape',
                    'Causeway.', 'Causwa.', 'Cen.', 'Cent.', 'Center.', 'Centers.', 'Centr',
                    'Centre.', 'Cir.', 'Circ.', 'Circl.', 'Circle.', 'Circles.', 'Cirs.', 'Clb',
                    'Clf.', 'Clfs.', 'Cliff.', 'Cliffs.', 'Club.', 'Cmn.', 'Cmns.', 'Cmp.', 'Cnter',
                    'Cntr.', 'Cnyn.', 'Common.', 'Commons.', 'Cor.', 'Corner.', 'Corners.', 'Cors',
                    'Course.', 'Court.', 'Courts.', 'Cove.', 'Coves.', 'Cp.', 'Cpe.', 'Crcl.', 'Crcle',
                    'Creek.', 'Cres.', 'Crescent.', 'Crest.', 'Crk.', 'Crossing.', 'Crossroad',
                    'Crossroads.', 'Crse.', 'Crsent.', 'Crsnt.', 'Crssng.', 'Crst.', 'Cswy.', 'Ct',
                    'Ctr.', 'Ctrs.', 'Cts.', 'Curv.', 'Curve.', 'Cv.', 'Cvs.', 'Cyn.', 'Dale.', 'Dam',
                    'Div.', 'Divide.', 'Dl.', 'Dm.', 'Dr.', 'Driv.', 'Drive.', 'Drives.', 'Drs.', 'Drv',
                    'Dv.', 'Dvd.', 'Est.', 'Estate.', 'Estates.', 'Ests.', 'Exp.', 'Expr.', 'Express',
                    'Expressway.', 'Expw.', 'Expy.', 'Ext.', 'Extension.', 'Extensions.', 'Extn',
                    'Extnsn.', 'Exts.', 'Fall.', 'Falls.', 'Ferry.', 'Field.', 'Fields.', 'Flat',
                    'Flats.', 'Fld.', 'Flds.', 'Fls.', 'Flt.', 'Flts.', 'Ford.', 'Fords.', 'Forest',
                    'Forests.', 'Forg.', 'Forge.', 'Forges.', 'Fork.', 'Forks.', 'Fort.', 'Frd',
                    'Frds.', 'Freeway.', 'Freewy.', 'Frg.', 'Frgs.', 'Frk.', 'Frks.', 'Frry.', 'Frst',
                    'Frt.', 'Frway.', 'Frwy.', 'Fry.', 'Ft.', 'Fwy.', 'Garden.', 'Gardens.', 'Gardn',
                    'Gateway.', 'Gatewy.', 'Gatway.', 'Gdn.', 'Gdns.', 'Glen.', 'Glens.', 'Gln',
                    'Glns.', 'Grden.', 'Grdn.', 'Grdns.', 'Green.', 'Greens.', 'Grn.', 'Grns.', 'Grov',
                    'Grove.', 'Groves.', 'Grv.', 'Grvs.', 'Gtway.', 'Gtwy.', 'Harb.', 'Harbor',
                    'Harbors.', 'Harbr.', 'Haven.', 'Hbr.', 'Hbrs.', 'Heights.', 'Highway.', 'Highwy',
                    'Hill.', 'Hills.', 'Hiway.', 'Hiwy.', 'Hl.', 'Hllw.', 'Hls.', 'Hollow.', 'Hollows',
                    'Holw.', 'Holws.', 'Hrbor.', 'Ht.', 'Hts.', 'Hvn.', 'Hway.', 'Hwy.', 'Inlet',
                    'Inlt.', 'Is.', 'Island.', 'Islands.', 'Isle.', 'Isles.', 'Islnd.', 'Islnds',
                    'Iss.', 'Jct.', 'Jction.', 'Jctn.', 'Jctns.', 'Jcts.', 'Junction.', 'Junctions',
                    'Junctn.', 'Juncton.', 'Key.', 'Keys.', 'Knl.', 'Knls.', 'Knol.', 'Knoll',
                    'Knolls.', 'Ky.', 'Kys.', 'Lake.', 'Lakes.', 'Land.', 'Landing.', 'Lane.', 'Lck',
                    'Lcks.', 'Ldg.', 'Ldge.', 'Lf.', 'Lgt.', 'Lgts.', 'Light.', 'Lights.', 'Lk.', 'Lks',
                    'Ln.', 'Lndg.', 'Lndng.', 'Loaf.', 'Lock.', 'Locks.', 'Lodg.', 'Lodge.', 'Loop',
                    'Loops.', 'Mall.', 'Manor.', 'Manors.', 'Mdw.', 'Mdws.', 'Meadow.', 'Meadows',
                    'Medows.', 'Mews.', 'Mill.', 'Mills.', 'Mission.', 'Missn.', 'Ml.', 'Mls.', 'Mnr',
                    'Mnrs.', 'Mnt.', 'Mntain.', 'Mntn.', 'Mntns.', 'Motorway.', 'Mount.', 'Mountain',
                    'Mountains.', 'Mountin.', 'Msn.', 'Mssn.', 'Mt.', 'Mtin.', 'Mtn.', 'Mtns.', 'Mtwy',
                    'Nck.', 'Neck.', 'Opas.', 'Orch.', 'Orchard.', 'Orchrd.', 'Oval.', 'Overpass',
                    'Ovl.', 'Park.', 'Parks.', 'Parkway.', 'Parkways.', 'Parkwy.', 'Pass.', 'Passage',
                    'Path.', 'Paths.', 'Pike.', 'Pikes.', 'Pine.', 'Pines.', 'Pkway.', 'Pkwy',
                    'Pkwys.', 'Pky.', 'Pl.', 'Place.', 'Plain.', 'Plains.', 'Plaza.', 'Pln.', 'Plns',
                    'Plz.', 'Plza.', 'Pne.', 'Pnes.', 'Point.', 'Points.', 'Port.', 'Ports.', 'Pr',
                    'Prairie.', 'Prk.', 'Prr.', 'Prt.', 'Prts.', 'Psge.', 'Pt.', 'Pts.', 'Rad',
                    'Radial.', 'Radiel.', 'Radl.', 'Ramp.', 'Ranch.', 'Ranches.', 'Rapid.', 'Rapids',
                    'Rd.', 'Rdg.', 'Rdge.', 'Rdgs.', 'Rds.', 'Rest.', 'Ridge.', 'Ridges.', 'Riv',
                    'River.', 'Rivr.', 'Rnch.', 'Rnchs.', 'Road.', 'Roads.', 'Route.', 'Row.', 'Rpd',
                    'Rpds.', 'Rst.', 'Rte.', 'Rue.', 'Run.', 'Rvr.', 'Shl.', 'Shls.', 'Shoal',
                    'Shoals.', 'Shoar.', 'Shoars.', 'Shore.', 'Shores.', 'Shr.', 'Shrs.', 'Skwy',
                    'Skyway.', 'Smt.', 'Spg.', 'Spgs.', 'Spng.', 'Spngs.', 'Spring.', 'Springs',
                    'Sprng.', 'Sprngs.', 'Spur.', 'Spurs.', 'Sq.', 'Sqr.', 'Sqre.', 'Sqrs.', 'Sqs',
                    'Squ.', 'Square.', 'Squares.', 'St.', 'Sta.', 'Station.', 'Statn.', 'Stn.', 'Str',
                    'Stra.', 'Strav.', 'Straven.', 'Stravenue.', 'Stravn.', 'Stream.', 'Street',
                    'Streets.', 'Streme.', 'Strm.', 'Strt.', 'Strvn.', 'Strvnue.', 'Sts.', 'Sumit',
                    'Sumitt.', 'Summit.', 'Ter.', 'Terr.', 'Terrace.', 'Throughway.', 'Tpke',
                    'Trace.', 'Traces.', 'Track.', 'Tracks.', 'Trafficway.', 'Trail.', 'Trailer',
                    'Trails.', 'Trak.', 'Trce.', 'Trfy.', 'Trk.', 'Trks.', 'Trl.', 'Trlr.', 'Trlrs',
                    'Trls.', 'Trnpk.', 'Trwy.', 'Tunel.', 'Tunl.', 'Tunls.', 'Tunnel.', 'Tunnels',
                    'Tunnl.', 'Turnpike.', 'Turnpk.', 'Un.', 'Underpass.', 'Union.', 'Unions.', 'Uns',
                    'Upas.', 'Valley.', 'Valleys.', 'Vally.', 'Vdct.', 'Via.', 'Viadct.', 'Viaduct',
                    'View.', 'Views.', 'Vill.', 'Villag.', 'Village.', 'Villages.', 'Ville.', 'Villg',
                    'Villiage.', 'Vis.', 'Vist.', 'Vista.', 'Vl.', 'Vlg.', 'Vlgs.', 'Vlly.', 'Vly',
                    'Vlys.', 'Vst.', 'Vsta.', 'Vw.', 'Vws.', 'Walk.', 'Walks.', 'Wall.', 'Way.', 'Ways',
                    'Well.', 'Wells.', 'Wl.', 'Wls.', 'Wy.', 'Xing.', 'Xrd.', 'Xrds',
                    'Allee', 'Alley', 'Ally', 'Aly', 'Anex', 'Annex', 'Annx', 'Anx', 'Arc',
                    'Arcade', 'Av', 'Ave', 'Aven', 'Avenu', 'Avenue', 'Avn', 'Avnue', 'Bayoo',
                    'Bayou', 'Bch', 'Beach', 'Bend', 'Bg', 'Bgs', 'Blf', 'Blfs', 'Bluf',
                    'Bluff', 'Bluffs', 'Blvd', 'Bnd', 'Bot', 'Bottm', 'Bottom', 'Boul',
                    'Boulevard', 'Boulv', 'Br', 'Branch', 'Brdge', 'Brg', 'Bridge', 'Brk',
                    'Brks', 'Brnch', 'Brook', 'Brooks', 'Btm', 'Burg', 'Burgs', 'Byp', 'Bypa',
                    'Bypas', 'Bypass', 'Byps', 'Byu', 'Camp', 'Canyn', 'Canyon', 'Cape',
                    'Causeway', 'Causwa', 'Cen', 'Cent', 'Center', 'Centers', 'Centr',
                    'Centre', 'Cir', 'Circ', 'Circl', 'Circle', 'Circles', 'Cirs', 'Clb',
                    'Clf', 'Clfs', 'Cliff', 'Cliffs', 'Club', 'Cmn', 'Cmns', 'Cmp', 'Cnter',
                    'Cntr', 'Cnyn', 'Common', 'Commons', 'Cor', 'Corner', 'Corners', 'Cors',
                    'Course', 'Court', 'Courts', 'Cove', 'Coves', 'Cp', 'Cpe', 'Crcl', 'Crcle',
                    'Creek', 'Cres', 'Crescent', 'Crest', 'Crk', 'Crossing', 'Crossroad',
                    'Crossroads', 'Crse', 'Crsent', 'Crsnt', 'Crssng', 'Crst', 'Cswy', 'Ct',
                    'Ctr', 'Ctrs', 'Cts', 'Curv', 'Curve', 'Cv', 'Cvs', 'Cyn', 'Dale', 'Dam',
                    'Div', 'Divide', 'Dl', 'Dm', 'Dr', 'Driv', 'Drive', 'Drives', 'Drs', 'Drv',
                    'Dv', 'Dvd', 'Est', 'Estate', 'Estates', 'Ests', 'Exp', 'Expr', 'Express',
                    'Expressway', 'Expw', 'Expy', 'Ext', 'Extension', 'Extensions', 'Extn',
                    'Extnsn', 'Exts', 'Fall', 'Falls', 'Ferry', 'Field', 'Fields', 'Flat',
                    'Flats', 'Fld', 'Flds', 'Fls', 'Flt', 'Flts', 'Ford', 'Fords', 'Forest',
                    'Forests', 'Forg', 'Forge', 'Forges', 'Fork', 'Forks', 'Fort', 'Frd',
                    'Frds', 'Freeway', 'Freewy', 'Frg', 'Frgs', 'Frk', 'Frks', 'Frry', 'Frst',
                    'Frt', 'Frway', 'Frwy', 'Fry', 'Ft', 'Fwy', 'Garden', 'Gardens', 'Gardn',
                    'Gateway', 'Gatewy', 'Gatway', 'Gdn', 'Gdns', 'Glen', 'Glens', 'Gln',
                    'Glns', 'Grden', 'Grdn', 'Grdns', 'Green', 'Greens', 'Grn', 'Grns', 'Grov',
                    'Grove', 'Groves', 'Grv', 'Grvs', 'Gtway', 'Gtwy', 'Harb', 'Harbor',
                    'Harbors', 'Harbr', 'Haven', 'Hbr', 'Hbrs', 'Heights', 'Highway', 'Highwy',
                    'Hill', 'Hills', 'Hiway', 'Hiwy', 'Hl', 'Hllw', 'Hls', 'Hollow', 'Hollows',
                    'Holw', 'Holws', 'Hrbor', 'Ht', 'Hts', 'Hvn', 'Hway', 'Hwy', 'Inlet',
                    'Inlt', 'Is', 'Island', 'Islands', 'Isle', 'Isles', 'Islnd', 'Islnds',
                    'Iss', 'Jct', 'Jction', 'Jctn', 'Jctns', 'Jcts', 'Junction', 'Junctions',
                    'Junctn', 'Juncton', 'Key', 'Keys', 'Knl', 'Knls', 'Knol', 'Knoll',
                    'Knolls', 'Ky', 'Kys', 'Lake', 'Lakes', 'Land', 'Landing', 'Lane', 'Lck',
                    'Lcks', 'Ldg', 'Ldge', 'Lf', 'Lgt', 'Lgts', 'Light', 'Lights', 'Lk', 'Lks',
                    'Ln', 'Lndg', 'Lndng', 'Loaf', 'Lock', 'Locks', 'Lodg', 'Lodge', 'Loop',
                    'Loops', 'Mall', 'Manor', 'Manors', 'Mdw', 'Mdws', 'Meadow', 'Meadows',
                    'Medows', 'Mews', 'Mill', 'Mills', 'Mission', 'Missn', 'Ml', 'Mls', 'Mnr',
                    'Mnrs', 'Mnt', 'Mntain', 'Mntn', 'Mntns', 'Motorway', 'Mount', 'Mountain',
                    'Mountains', 'Mountin', 'Msn', 'Mssn', 'Mt', 'Mtin', 'Mtn', 'Mtns', 'Mtwy',
                    'Nck', 'Neck', 'Opas', 'Orch', 'Orchard', 'Orchrd', 'Oval', 'Overpass',
                    'Ovl', 'Park', 'Parks', 'Parkway', 'Parkways', 'Parkwy', 'Pass', 'Passage',
                    'Path', 'Paths', 'Pike', 'Pikes', 'Pine', 'Pines', 'Pkway', 'Pkwy',
                    'Pkwys', 'Pky', 'Pl', 'Place', 'Plain', 'Plains', 'Plaza', 'Pln', 'Plns',
                    'Plz', 'Plza', 'Pne', 'Pnes', 'Point', 'Points', 'Port', 'Ports', 'Pr',
                    'Prairie', 'Prk', 'Prr', 'Prt', 'Prts', 'Psge', 'Pt', 'Pts', 'Rad',
                    'Radial', 'Radiel', 'Radl', 'Ramp', 'Ranch', 'Ranches', 'Rapid', 'Rapids',
                    'Rd', 'Rdg', 'Rdge', 'Rdgs', 'Rds', 'Rest', 'Ridge', 'Ridges', 'Riv',
                    'River', 'Rivr', 'Rnch', 'Rnchs', 'Road', 'Roads', 'Route', 'Row', 'Rpd',
                    'Rpds', 'Rst', 'Rte', 'Rue', 'Run', 'Rvr', 'Shl', 'Shls', 'Shoal',
                    'Shoals', 'Shoar', 'Shoars', 'Shore', 'Shores', 'Shr', 'Shrs', 'Skwy',
                    'Skyway', 'Smt', 'Spg', 'Spgs', 'Spng', 'Spngs', 'Spring', 'Springs',
                    'Sprng', 'Sprngs', 'Spur', 'Spurs', 'Sq', 'Sqr', 'Sqre', 'Sqrs', 'Sqs',
                    'Squ', 'Square', 'Squares', 'St', 'Sta', 'Station', 'Statn', 'Stn', 'Str',
                    'Stra', 'Strav', 'Straven', 'Stravenue', 'Stravn', 'Stream', 'Street',
                    'Streets', 'Streme', 'Strm', 'Strt', 'Strvn', 'Strvnue', 'Sts', 'Sumit',
                    'Sumitt', 'Summit', 'Ter', 'Terr', 'Terrace', 'Throughway', 'Tpke',
                    'Trace', 'Traces', 'Track', 'Tracks', 'Trafficway', 'Trail', 'Trailer',
                    'Trails', 'Trak', 'Trce', 'Trfy', 'Trk', 'Trks', 'Trl', 'Trlr', 'Trlrs',
                    'Trls', 'Trnpk', 'Trwy', 'Tunel', 'Tunl', 'Tunls', 'Tunnel', 'Tunnels',
                    'Tunnl', 'Turnpike', 'Turnpk', 'Un', 'Underpass', 'Union', 'Unions', 'Uns',
                    'Upas', 'Valley', 'Valleys', 'Vally', 'Vdct', 'Via', 'Viadct', 'Viaduct',
                    'View', 'Views', 'Vill', 'Villag', 'Village', 'Villages', 'Ville', 'Villg',
                    'Villiage', 'Vis', 'Vist', 'Vista', 'Vl', 'Vlg', 'Vlgs', 'Vlly', 'Vly',
                    'Vlys', 'Vst', 'Vsta', 'Vw', 'Vws', 'Walk', 'Walks', 'Wall', 'Way', 'Ways',
                    'Well', 'Wells', 'Wl', 'Wls', 'Wy', 'Xing', 'Xrd', 'Xrds',
                    }
    try:
        address_data_tag = usaddress.tag(full_address, tag_mapping={
            'AddressNumber': 'address1',
            'AddressNumberPrefix': 'address1',
            'AddressNumberSuffix': 'address1',
            'StreetName': 'address1',
            'StreetNamePreDirectional': 'address1',
            'StreetNamePreModifier': 'address1',
            'StreetNamePreType': 'address1',
            'StreetNamePostDirectional': 'address1',
            'StreetNamePostModifier': 'address1',
            'StreetNamePostType': 'address1',
            'CornerOf': 'address1',
            'IntersectionSeparator': 'address1',
            'LandmarkName': 'address1',
            'USPSBoxGroupID': 'address1',
            'USPSBoxGroupType': 'address1',
            'USPSBoxID': 'address1',
            'USPSBoxType': 'address1',
            'OccupancyType': 'address2',
            'OccupancyIdentifier': 'address2',
            'PlaceName': 'city',
            'StateName': 'state',
            'ZipCode': 'zip_code',
        })
        address_data = address_data_tag[0]
        is_address_data_tagged = True
    except Exception as ae:
        is_address_data_tagged = False
        error_message = str(ae)

    if is_address_data_tagged:
        addressLine1 = address_data['address1'] if 'address1' in address_data else ''
        addressLine2 = address_data['address2'] if 'address2' in address_data else ''
        addressCity = address_data['city'] if 'city' in address_data else ''
        addressState = address_data['state'] if 'state' in address_data else ''
        addressZip = address_data['zip_code'] if 'zip_code' in address_data else ''

        if addressLine1 == '' or addressCity == '' or addressState == '' or addressZip == '':
            addressLine1 = ''
            addressLine2 = ''
            addressCity = ''
            addressState = ''
            addressZip = ''

        if ' ' in addressCity:
            space = addressCity.index(' ')
            test_for_address_line_one_word = addressCity.rsplit('space')
            if test_for_address_line_one_word[0] in street_names:
                addressLine1 = ''
                addressLine2 = ''
                addressCity = ''
                addressState = ''
                addressZip = ''

        if len(addressCity) > 0 and ' ' in addressCity:
            city_str = ''
            vals = addressCity.split()
            for val in vals:
                if val in street_names or 'APT' in val:
                    addressLine2 = addressLine2 + ' ' + val
                else:
                    if len(city_str) > 0:
                        city_str = city_str + ' ' + val
                    else:
                        city_str = val

            addressCity = city_str
    else:
        addressLine1 = ''
        addressLine2 = ''
        addressCity = ''
        addressState = ''
        addressZip = ''

    return addressLine1, addressLine2, addressCity, addressState, addressZip


def remove_non_digit(s):
    input_str = str(s)
    if len(input_str) > 0:
        return re.sub('\D', '', input_str)
    else:
        return None


def current_time_min_str():
    return datetime.now().strftime('%Y%m%d%H%M')


def parse_name(nameStr):
    full_names = None
    if len(nameStr) > 0 and ' ' in nameStr:
        full_names = Names()
        name_str = None
        if ',' not in nameStr and ' ' in nameStr:
            name_parts = str(nameStr).split()
            part1 = name_parts[0]
            part12 = ' '.join(name_parts[:2])
            if part12.upper() in lastNameMultipartStrings:
                name_str = part12 + ' ' + name_parts[2] + ', ' + ' '.join(name_parts[3:])
            elif part1.upper() in lastNameMultipartStrings:
                name_str = part1 + ' ' + name_parts[1] + ', ' + ' '.join(name_parts[2:])
            else:
                suffix_list = suffix_dictonary.keys()
                delimiter = 0
                for index in range(0, len(name_parts)):
                    if name_parts[index] in suffix_list:
                        delimiter = index
                        break

                if delimiter > 0:
                    name_str = ' '.join(name_parts[:(delimiter+1)]) + ', ' + ' '.join(name_parts[(delimiter+1):])
                else:
                    name_str = name_parts[0] + ', ' + ' '.join(name_parts[1:])

        if name_str is None:
            name_str = nameStr

        parsed_name = HumanName(name_str.encode('utf-8'))
        parsed_full_name = parsed_name.full_name
        if parsed_name.suffix is not None:
            full_names.suffix = parsed_name.suffix.title()
        if parsed_name.first is not None:
            full_names.first = parsed_name.first.title()
        if parsed_name.last is not None:
            full_names.last = parsed_name.last.title()
        if parsed_name.middle is not None:
            full_names.middle = parsed_name.middle.title()

        name_parts = str(parsed_full_name).split(',')
        last_name_str = name_parts[0].strip()
        first_name_str = name_parts[1].strip()
        first_name_parts = first_name_str.split(' ')

        if len(first_name_parts) > 2:
            full_names.first = ' '.join(first_name_parts[:2])
            full_names.middle = ' '.join(first_name_parts[2:])

        if len(parsed_name.last) == 0 or len(
                parsed_name.first) == 0 or parsed_name.last.upper() not in last_name_str.upper():
            full_names.first = first_name_str
            full_names.last = last_name_str
            full_names.middle = ''

    return full_names


def name_parser(nameStr):
    full_names = Names()
    name_str = nameStr

    if ',' not in nameStr and ' ' in nameStr:
        name_parts = str(nameStr).split()
        part1 = name_parts[0]
        part12 = ' '.join(name_parts[:2])
        if part12.upper() in lastNameMultipartStrings:
            name_str = part12 + ' ' + name_parts[2] + ', ' + ' '.join(name_parts[3:])
        elif part1.upper() in lastNameMultipartStrings:
            name_str = part1 + ' ' + name_parts[1] + ', ' + ' '.join(name_parts[2:])
        else:
            suffix_list = suffix_dictonary.keys()
            delimiter = 0
            for index in range(0, len(name_parts)):
                if name_parts[index] in suffix_list:
                    delimiter = index
                    break

            if delimiter > 0:
                name_str = ' '.join(name_parts[:(delimiter + 1)]) + ', ' + ' '.join(name_parts[(delimiter + 1):])
            else:
                name_str = name_parts[0] + ', ' + ' '.join(name_parts[1:])

    if ',' in name_str:
        name_items = name_str.split(',')
        full_names.last = name_items[0].strip()
        full_names.first = name_items[1].strip()
    else:
        full_names.last = name_str

    if ' ' in full_names.first:
        first_name_parts = full_names.first.split()
        full_names.first = first_name_parts[0]
        full_names.middle = ' '.join(first_name_parts[1:])

    if ' ' in full_names.last:
        last_name_parts = full_names.last.split()
        suffix_list = suffix_dictonary.keys()
        delimiter = 0
        for index in range(0, len(last_name_parts)):
            if last_name_parts[index] in suffix_list:
                delimiter = index
                break

        if delimiter > 0:
            full_names.last = ' '.join(last_name_parts[:delimiter])
            full_names.suffix = last_name_parts[delimiter]

    return full_names


def parse_city_state_zip(inputStr):
    city = ''
    state = ''
    zip = ''

    if len(inputStr) > 0 and ' ' in inputStr:
        vals = inputStr.split()
        delimiter = 10
        for index in reversed(range(len(vals))):
            if vals[index] in us_states:
                delimiter = index
                state = vals[index]
                break

        if delimiter < 10:
            city = ' '.join(vals[:delimiter])
            if delimiter < len(vals):
                zip = ' '.join(vals[delimiter+1:])
        else:
            cnt = len(vals)
            if cnt >= 3:
                zip = vals[cnt-1]
                state = vals[cnt-2]
                city = ' '.join(vals[:(cnt-2)])
            else:
                if (vals[1]).isdigit:
                    zip = vals[1]
                    if len(vals[0]) == 2:
                        state = vals[0]
                    else:
                        city = vals[0]

    return city, state, zip


def between_x1(cursor, start_x1, end_x1, row):
    VALUES = cursor.execute(
        "SELECT ElementText from pageelement where X1 >= ? and X1 < ? and ElementRow = ?",
        (start_x1, end_x1, row,)).fetchall()
    if VALUES is None or len(VALUES) == 0:
        return ""

    if len(VALUES) == 1:
        return VALUES[0][0]

    return " ".join(zip(*VALUES)[0])


def between_columns(cursor, start_column, end_column, row):
    VALUES = cursor.execute(
        "SELECT ElementText from pageelement where ElementColumn >= ? and ElementColumn < ? and ElementRow = ?",
        (start_column, end_column, row,)).fetchall()
    if VALUES is None or len(VALUES) == 0:
        return ""

    if len(VALUES) == 1:
        return VALUES[0][0]

    return " ".join(zip(*VALUES)[0])


def between_rows(cursor, start_row, end_row, column):
    VALUES = cursor.execute(
        "SELECT ElementText from pageelement where ElementRow >= ? and ElementRow <= ? and ElementColumn = ?",
        (start_row, end_row, column,)).fetchall()
    if VALUES is None or len(VALUES) == 0:
        return ""

    if len(VALUES) == 1:
        return VALUES[0][0]

    return " ".join(zip(*VALUES)[0])


def between_text(cursor, start_line, end_line, start_text, end_text):
    startColumn = cursor.execute(
        "SELECT ElementColumn, ElementRow from pageelement where ElementRow >= ? and ElementRow <= ? and ElementText "
        "= ?",
        (start_line, end_line, start_text,)).fetchone()
    if startColumn is None:
        return ""
    endColumn = cursor.execute(
        "SELECT ElementColumn from pageelement where ElementRow >= ? and ElementRow <= ? and ElementText = ?",
        (start_line, end_line, end_text,)).fetchone()
    if endColumn is None:
        return ""
    VALUES = cursor.execute(
        "SELECT elementText from pageelement where ElementColumn > ? and ElementColumn < ? and ElementRow = ?",
        (startColumn[0], endColumn[0], startColumn[1],)).fetchall()
    if VALUES is None or len(VALUES) == 0:
        return ""
    return " ".join(zip(*VALUES)[0])


def next_text(cursor, start_line, end_line, start_text):
    startColumn = cursor.execute(
        "SELECT ElementColumn, ElementRow, X1 from pageelement where ElementRow >= ? and ElementRow <= ? and "
        "ElementText = ?",
        (start_line, end_line, start_text,)).fetchone()
    if startColumn is None:
        return ""
    VALUES = cursor.execute(
        "SELECT elementText from pageelement where ElementColumn >= ? and ElementRow = ? and X1 > ? Order by "
        "ElementColumn, X1 LIMIT 1",
        (startColumn[0], startColumn[1], startColumn[2],)).fetchone()
    if VALUES is None or len(VALUES) == 0:
        return ""

    return VALUES[0]


def next_text_before_x(cursor, start_line, end_line, start_text, end_x):
    startColumn = cursor.execute(
        "SELECT ElementColumn, ElementRow, X1 from pageelement where ElementRow >= ? and ElementRow <= ? and "
        "ElementText like ?",
        (start_line, end_line, start_text + '%',)).fetchone()
    if startColumn is None:
        return ""
    VALUES = cursor.execute(
        "SELECT elementText from pageelement where ElementRow = ? and X1 >= ? and x1 < ? Order by "
        "ElementColumn, X1",
        (startColumn[1], startColumn[2] - 50, end_x,)).fetchall()
    if VALUES is None or len(VALUES) == 0:
        return ""

    value_text = "".join(x[0] for x in VALUES)
    value_text = value_text.replace(start_text, "").strip()

    return value_text


def get_effective_date():
    today = datetime.now().date()
    this_year = today.year

    if today.month >= 3 and today.day >= 1:
        effectiveDates_effectiveDate1 = datetime.strptime('1 Jan ' + str(this_year), '%d %b %Y')
    else:
        effectiveDates_effectiveDate1 = datetime.strptime('1 Jan ' + str(this_year - 1), '%d %b %Y')

    return effectiveDates_effectiveDate1


# The following headers are not matched to EmployeeComplex object yet:
# employee_headers = ['action', 'clientId', 'employeeNumber', 'prefix', 'accredited',
#                    'terminationReason', 'flsa', 'annualHours', 'ownerOfficer', 'baseShift',
#					'adjustedHireDate', 'seniorityDate', 'checkPrintSort', 'managerClientId']
def generate_model_employee(employee):
    model_employee = None
    emp_manager = None
    emp_keys = employee.keys()
    if len(emp_keys) > 0:
        model_employee = EmployeeComplex()
        full_names = Names()
        if 'lastName' in emp_keys:
            full_names.last = employee['lastName']
            emp_keys.remove('lastName')
        if 'firstName' in emp_keys:
            full_names.first = employee['firstName']
            emp_keys.remove('firstName')
        if 'middleName' in emp_keys:
            full_names.middle = employee['middleName']
            emp_keys.remove('middleName')
        if 'suffix' in emp_keys:
            full_names.suffix = get_lookup_value(employee['suffix'], suffixLookup)
            emp_keys.remove('suffix')

        model_employee.names = full_names

        emp_keys.remove('priorCompanyCode')

        for key in emp_keys:
            val = employee[key]
            if key == 'priorEmployeeNumber':
                model_employee.employeeId = val
            elif key == 'ssn':
                model_employee.ssn = val
            # elif key == '':
            #     model_employee.paycorNumber = val
            elif key == 'departmentCode':
                model_employee.dept = val
            elif key == 'departmentDescription':
                model_employee.deptDescription = val
            # elif key == '':
            #     model_employee.status = val
            elif key == 'payrollCode':
                model_employee.payrollCode = val
            elif key == 'paygroupDescription':
                model_employee.priorPaygroup = val
            elif key == 'employmentStatus':
                model_employee.statusCode = get_lookup_value(val, employmentStatus)
            elif key == 'hireDate':
                model_employee.hireDate = val
            elif key == 'reHireDate':
                model_employee.rehireDate = val
            elif key == 'birthDate':
                model_employee.birthDate = val
            elif key == 'gender' or key == 'sex':
                model_employee.sex = val
            elif key == 'terminationDate':
                model_employee.termDate = val
            elif key == 'employeeType':
                model_employee.employeeType = val
            elif key == 'ethnicity':
                model_employee.race = val
            elif key == 'maritalStatus':
                model_employee.maritalStatus = val
            # elif key == '':
            #     model_employee.country = val
            elif key == 'state':
                model_employee.state = val
            elif key == 'zip':
                model_employee.zip = val
            elif key == 'city':
                model_employee.city = val
            elif key == 'addressLine1':
                model_employee.address1 = val
            elif key == 'addressLine2':
                model_employee.address2 = val
            elif key == 'statusType':
                model_employee.employmentStatusType = get_lookup_value(val, statusLookup)
            elif key == 'jobTitle' or key == 'jobTitleCode':
                model_employee.jobTitle = val
            elif key == 'mobilePhone':
                model_employee.phoneMobile = val
            elif key == 'homePhone':
                model_employee.phoneHome = val
            elif key == 'workPhone':
                model_employee.phoneWork = val
            elif key == 'workPhoneNumberExtension':
                model_employee.phoneWorkExtension = val
            elif key == 'homeEmail':
                model_employee.emailHome = val
            elif key == 'workEmail':
                model_employee.emailWork = val
            elif 'manager' in key:
                if emp_manager is None:
                    emp_manager = Manager()
                if key == 'managerEmployeeNumber':
                    emp_manager.employeeId = val
                elif key == 'managerLastName':
                    emp_manager.lastName = val
                elif key == 'managerFirstName':
                    emp_manager.firstName = val
            else:
                keyVal = str(key)
                keyVal = keyVal.replace(':', '')
                model_employee.unknownValues.update({keyVal: val})

        if emp_manager is not None:
            model_employee.manager = emp_manager

    return model_employee


def create_clientCode(record):
    if len(record) > 0:
        clientCode_record = ClientCode()
        for key in record.keys():
            val = record[key]
            if key == 'code' or key == 'taxCode':
                clientCode_record.code = val
            if key == 'altCodes':
                clientCode_record.altCodes = val
            if key == 'effectiveDates amount1':
                clientCode_record.amount = val
            if key == 'effectiveDates rate1':
                clientCode_record.rate = val
            if key == 'effectiveDates effectiveDate1' or key == 'effectiveDate':
                clientCode_record.effectiveDate = val
            if key == 'calculate':
                clientCode_record.calculate = val

        return clientCode_record


