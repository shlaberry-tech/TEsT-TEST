# Reconstructed from the supplied executable; not the original source.
# Requested SME product grouping and two MAZIWA sheets added.
# Decompiled keyword calls and the damaged final sizing/save block repaired.
# Other recovered logic is retained and needs validation against real reports.
# Exact, case-sensitive product names are used. Bare "SME" is excluded.
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.formatting.rule import IconSetRule, Rule, FormatObject, IconSet, ColorScale
from datetime import datetime, timedelta
from pandas.tseries.offsets import MonthEnd
from tqdm import tqdm
import os
import re
from openpyxl.utils import get_column_letter
from openpyxl.styles import Border, Side, Color
import calendar
SME_PRODUCTS = ('SME Commercial', 'SME Quick', 'SME Secured', 'Business Improvement')
delta_icon = '▲'
day_month_year = datetime.today().strftime('%Y_%m_%d')

def set_border(ws, cell_range):
    thin = Side(border_style='thin', color='000000')
    for row in ws[cell_range]:
        for cell in row:
            cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)

def summary_headings(summary_sheet):
    summary_sheet['A1'] = 'Product'
    summary_sheet['A1'].font = Font(size=10, bold=True)
    summary_sheet['A1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    summary_sheet['B1'] = 'Book Size'
    summary_sheet['B1'].font = Font(size=10, bold=True)
    summary_sheet['B1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    summary_sheet['C1'] = 'VAR 30'
    summary_sheet['C1'].font = Font(size=10, bold=True)
    summary_sheet['C1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    summary_sheet['D1'] = 'PAR 30'
    summary_sheet['D1'].font = Font(size=10, bold=True)
    summary_sheet['D1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    summary_sheet['E1'] = delta_icon + ' from Yesterday(VAR)'
    summary_sheet['E1'].font = Font(size=10, bold=True)
    summary_sheet['E1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    summary_sheet['F1'] = delta_icon + ' from Beg. of Month(VAR)'
    summary_sheet['F1'].font = Font(size=10, bold=True)
    summary_sheet['F1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    summary_sheet['G1'] = delta_icon + ' from Beg. of Year(VAR)'
    summary_sheet['G1'].font = Font(size=10, bold=True)
    summary_sheet['G1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    summary_sheet['H1'] = delta_icon + ' from Beg. of Month(Book Size)'
    summary_sheet['H1'].font = Font(size=10, bold=True)
    summary_sheet['H1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    summary_sheet['I1'] = delta_icon + ' from Beg. of Year(Book Size)'
    summary_sheet['I1'].font = Font(size=10, bold=True)
    summary_sheet['I1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')

def collections_headings(collections_sheet):
    collections_sheet['A1'] = 'Collection Officer'
    collections_sheet['A1'].font = Font(size=10, bold=True)
    collections_sheet['A1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    collections_sheet['B1'] = 'Book Size'
    collections_sheet['B1'].font = Font(size=10, bold=True)
    collections_sheet['B1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    collections_sheet['C1'] = 'VAR 30'
    collections_sheet['C1'].font = Font(size=10, bold=True)
    collections_sheet['C1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    collections_sheet['D1'] = 'PAR 30'
    collections_sheet['D1'].font = Font(size=10, bold=True)
    collections_sheet['D1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    collections_sheet['E1'] = 'Performing/ Normal'
    collections_sheet['E1'].font = Font(size=10, bold=True)
    collections_sheet['E1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    collections_sheet['F1'] = 'Partial Installment 1'
    collections_sheet['F1'].font = Font(size=10, bold=True)
    collections_sheet['F1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    collections_sheet['G1'] = 'Partial Installment 2'
    collections_sheet['G1'].font = Font(size=10, bold=True)
    collections_sheet['G1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    collections_sheet['H1'] = 'NPL-Vote'
    collections_sheet['H1'].font = Font(size=10, bold=True)
    collections_sheet['H1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    collections_sheet['I1'] = 'NPL-Client'
    collections_sheet['I1'].font = Font(size=10, bold=True)
    collections_sheet['I1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    collections_sheet['J1'] = 'Zero Receipted'
    collections_sheet['J1'].font = Font(size=10, bold=True)
    collections_sheet['J1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    collections_sheet['K1'] = 'VAR ' + delta_icon + ' from Yesterday'
    collections_sheet['K1'].font = Font(size=10, bold=True)
    collections_sheet['K1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    collections_sheet['L1'] = 'VAR ' + delta_icon + ' from Beg. of the Month'
    collections_sheet['L1'].font = Font(size=10, bold=True)
    collections_sheet['L1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    collections_sheet['M1'] = 'VAR ' + delta_icon + ' from Beg. of the Year'
    collections_sheet['M1'].font = Font(size=10, bold=True)
    collections_sheet['M1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    collections_sheet['N1'] = 'Book Size ' + delta_icon + ' from Beg. of the Month'
    collections_sheet['N1'].font = Font(size=10, bold=True)
    collections_sheet['N1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    collections_sheet['O1'] = 'Book Size ' + delta_icon + ' from Beg. of the Year'
    collections_sheet['O1'].font = Font(size=10, bold=True)
    collections_sheet['O1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')

def cs_branches_summary_headings(branch_summary_sheet):
    branch_summary_sheet['A1'] = 'Branch'
    branch_summary_sheet['A1'].font = Font(size=10, bold=True)
    branch_summary_sheet['A1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    branch_summary_sheet['B1'] = 'Book Size'
    branch_summary_sheet['B1'].font = Font(size=10, bold=True)
    branch_summary_sheet['B1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    branch_summary_sheet['C1'] = 'VAR 30'
    branch_summary_sheet['C1'].font = Font(size=10, bold=True)
    branch_summary_sheet['C1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    branch_summary_sheet['D1'] = 'PAR 30'
    branch_summary_sheet['D1'].font = Font(size=10, bold=True)
    branch_summary_sheet['D1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    branch_summary_sheet['E1'] = 'BP PAR'
    branch_summary_sheet['E1'].font = Font(size=10, bold=True)
    branch_summary_sheet['E1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    branch_summary_sheet['F1'] = 'Performing/ Normal'
    branch_summary_sheet['F1'].font = Font(size=10, bold=True)
    branch_summary_sheet['F1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    branch_summary_sheet['G1'] = 'Partial Installment 1'
    branch_summary_sheet['G1'].font = Font(size=10, bold=True)
    branch_summary_sheet['G1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    branch_summary_sheet['H1'] = 'Partial Installment 2'
    branch_summary_sheet['H1'].font = Font(size=10, bold=True)
    branch_summary_sheet['H1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    branch_summary_sheet['I1'] = 'NPL-Vote'
    branch_summary_sheet['I1'].font = Font(size=10, bold=True)
    branch_summary_sheet['I1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    branch_summary_sheet['J1'] = 'NPL-Client'
    branch_summary_sheet['J1'].font = Font(size=10, bold=True)
    branch_summary_sheet['J1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    branch_summary_sheet['K1'] = 'Zero Receipted'
    branch_summary_sheet['K1'].font = Font(size=10, bold=True)
    branch_summary_sheet['K1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    branch_summary_sheet['L1'] = 'VAR ' + delta_icon + ' from Yest.'
    branch_summary_sheet['L1'].font = Font(size=10, bold=True)
    branch_summary_sheet['L1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    branch_summary_sheet['M1'] = 'VAR ' + delta_icon + ' from Beg. of Month'
    branch_summary_sheet['M1'].font = Font(size=10, bold=True)
    branch_summary_sheet['M1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')

def lbf_headings(lbf_sheet):
    lbf_sheet['A1'] = 'RO'
    lbf_sheet['A1'].font = Font(size=10, bold=True)
    lbf_sheet['A1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_sheet['B1'] = 'Book Size'
    lbf_sheet['B1'].font = Font(size=10, bold=True)
    lbf_sheet['B1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_sheet['C1'] = 'VAR 7'
    lbf_sheet['C1'].font = Font(size=10, bold=True)
    lbf_sheet['C1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_sheet['D1'] = 'PAR 7'
    lbf_sheet['D1'].font = Font(size=10, bold=True)
    lbf_sheet['D1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_sheet['E1'] = 'VAR 30'
    lbf_sheet['E1'].font = Font(size=10, bold=True)
    lbf_sheet['E1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_sheet['F1'] = 'PAR 30'
    lbf_sheet['F1'].font = Font(size=10, bold=True)
    lbf_sheet['F1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_sheet['G1'] = 'Repossession'
    lbf_sheet['G1'].font = Font(size=10, bold=True)
    lbf_sheet['G1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_sheet['H1'] = 'Yard'
    lbf_sheet['H1'].font = Font(size=10, bold=True)
    lbf_sheet['H1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_sheet['I1'] = 'Sold'
    lbf_sheet['I1'].font = Font(size=10, bold=True)
    lbf_sheet['I1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_sheet['J1'] = 'VAR30 ' + delta_icon + ' from Yest.'
    lbf_sheet['J1'].font = Font(size=10, bold=True)
    lbf_sheet['J1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_sheet['K1'] = 'VAR30 ' + delta_icon + ' from Beg. of Month'
    lbf_sheet['K1'].font = Font(size=10, bold=True)
    lbf_sheet['K1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_sheet['L1'] = 'VAR30 ' + delta_icon + ' from Beg. of Year'
    lbf_sheet['L1'].font = Font(size=10, bold=True)
    lbf_sheet['L1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_sheet['M1'] = 'Book Size ' + delta_icon + ' from Beg. of Month'
    lbf_sheet['M1'].font = Font(size=10, bold=True)
    lbf_sheet['M1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_sheet['N1'] = 'Book Size ' + delta_icon + ' from Beg. of Year'
    lbf_sheet['N1'].font = Font(size=10, bold=True)
    lbf_sheet['N1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')

def lbf_branches_headings(lbf_branch_sheet):
    lbf_branch_sheet['A1'] = 'LBF Branches'
    lbf_branch_sheet['A1'].font = Font(size=10, bold=True)
    lbf_branch_sheet['A1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_branch_sheet['B1'] = 'Book Size'
    lbf_branch_sheet['B1'].font = Font(size=10, bold=True)
    lbf_branch_sheet['B1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_branch_sheet['C1'] = 'VAR 7'
    lbf_branch_sheet['C1'].font = Font(size=10, bold=True)
    lbf_branch_sheet['C1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_branch_sheet['D1'] = 'PAR 7'
    lbf_branch_sheet['D1'].font = Font(size=10, bold=True)
    lbf_branch_sheet['D1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_branch_sheet['E1'] = 'VAR 30'
    lbf_branch_sheet['E1'].font = Font(size=10, bold=True)
    lbf_branch_sheet['E1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_branch_sheet['F1'] = 'PAR 30'
    lbf_branch_sheet['F1'].font = Font(size=10, bold=True)
    lbf_branch_sheet['F1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_branch_sheet['G1'] = 'Repossession'
    lbf_branch_sheet['G1'].font = Font(size=10, bold=True)
    lbf_branch_sheet['G1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_branch_sheet['H1'] = 'Yard'
    lbf_branch_sheet['H1'].font = Font(size=10, bold=True)
    lbf_branch_sheet['H1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_branch_sheet['I1'] = 'Sold'
    lbf_branch_sheet['I1'].font = Font(size=10, bold=True)
    lbf_branch_sheet['I1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_branch_sheet['J1'] = 'VAR30 ' + delta_icon + ' from Yest.'
    lbf_branch_sheet['J1'].font = Font(size=10, bold=True)
    lbf_branch_sheet['J1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_branch_sheet['K1'] = 'VAR30 ' + delta_icon + ' from Beg. of Month'
    lbf_branch_sheet['K1'].font = Font(size=10, bold=True)
    lbf_branch_sheet['K1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_branch_sheet['L1'] = 'VAR30 ' + delta_icon + ' from Beg. of Year'
    lbf_branch_sheet['L1'].font = Font(size=10, bold=True)
    lbf_branch_sheet['L1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_branch_sheet['M1'] = 'Book Size ' + delta_icon + ' from Beg. of Month'
    lbf_branch_sheet['M1'].font = Font(size=10, bold=True)
    lbf_branch_sheet['M1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_branch_sheet['N1'] = 'Book Size ' + delta_icon + ' from Beg. of Year'
    lbf_branch_sheet['N1'].font = Font(size=10, bold=True)
    lbf_branch_sheet['N1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')

def lbf_branches_headings(lbf_branch_sheet):
    lbf_branch_sheet['A1'] = 'LBF Branches'
    lbf_branch_sheet['A1'].font = Font(size=10, bold=True)
    lbf_branch_sheet['A1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_branch_sheet['B1'] = 'Book Size'
    lbf_branch_sheet['B1'].font = Font(size=10, bold=True)
    lbf_branch_sheet['B1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_branch_sheet['C1'] = 'VAR 7'
    lbf_branch_sheet['C1'].font = Font(size=10, bold=True)
    lbf_branch_sheet['C1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_branch_sheet['D1'] = 'PAR 7'
    lbf_branch_sheet['D1'].font = Font(size=10, bold=True)
    lbf_branch_sheet['D1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_branch_sheet['E1'] = 'VAR 30'
    lbf_branch_sheet['E1'].font = Font(size=10, bold=True)
    lbf_branch_sheet['E1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_branch_sheet['F1'] = 'PAR 30'
    lbf_branch_sheet['F1'].font = Font(size=10, bold=True)
    lbf_branch_sheet['F1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_branch_sheet['G1'] = 'Repossession'
    lbf_branch_sheet['G1'].font = Font(size=10, bold=True)
    lbf_branch_sheet['G1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_branch_sheet['H1'] = 'Yard'
    lbf_branch_sheet['H1'].font = Font(size=10, bold=True)
    lbf_branch_sheet['H1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_branch_sheet['I1'] = 'Sold'
    lbf_branch_sheet['I1'].font = Font(size=10, bold=True)
    lbf_branch_sheet['I1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_branch_sheet['J1'] = 'VAR30 ' + delta_icon + ' from Yest.'
    lbf_branch_sheet['J1'].font = Font(size=10, bold=True)
    lbf_branch_sheet['J1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_branch_sheet['K1'] = 'VAR30 ' + delta_icon + ' from Beg. of Month'
    lbf_branch_sheet['K1'].font = Font(size=10, bold=True)
    lbf_branch_sheet['K1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_branch_sheet['L1'] = 'VAR30 ' + delta_icon + ' from Beg. of Year'
    lbf_branch_sheet['L1'].font = Font(size=10, bold=True)
    lbf_branch_sheet['L1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_branch_sheet['M1'] = 'Book Size ' + delta_icon + ' from Beg. of Month'
    lbf_branch_sheet['M1'].font = Font(size=10, bold=True)
    lbf_branch_sheet['M1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    lbf_branch_sheet['N1'] = 'Book Size ' + delta_icon + ' from Beg. of Year'
    lbf_branch_sheet['N1'].font = Font(size=10, bold=True)
    lbf_branch_sheet['N1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')

def private_sector_employer_headings(private_sector_employer_sheet):
    private_sector_employer_sheet['A1'] = 'Employer'
    private_sector_employer_sheet['A1'].font = Font(size=10, bold=True)
    private_sector_employer_sheet['A1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    private_sector_employer_sheet['B1'] = 'Book Size'
    private_sector_employer_sheet['B1'].font = Font(size=10, bold=True)
    private_sector_employer_sheet['B1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    private_sector_employer_sheet['C1'] = 'VAR 30'
    private_sector_employer_sheet['C1'].font = Font(size=10, bold=True)
    private_sector_employer_sheet['C1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    private_sector_employer_sheet['D1'] = 'PAR 30'
    private_sector_employer_sheet['D1'].font = Font(size=10, bold=True)
    private_sector_employer_sheet['D1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    private_sector_employer_sheet['E1'] = 'VAR30 ' + delta_icon + ' from Beg. of Month'
    private_sector_employer_sheet['E1'].font = Font(size=10, bold=True)
    private_sector_employer_sheet['E1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    private_sector_employer_sheet['F1'] = 'VAR30 ' + delta_icon + ' from Beg. of Year'
    private_sector_employer_sheet['F1'].font = Font(size=10, bold=True)
    private_sector_employer_sheet['F1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    private_sector_employer_sheet['G1'] = 'Book Size ' + delta_icon + ' from Beg. of Month'
    private_sector_employer_sheet['G1'].font = Font(size=10, bold=True)
    private_sector_employer_sheet['G1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    private_sector_employer_sheet['H1'] = 'Book Size ' + delta_icon + ' from Beg. of Year'
    private_sector_employer_sheet['H1'].font = Font(size=10, bold=True)
    private_sector_employer_sheet['H1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')

def sme_headings(sme_sheet):
    sme_sheet['A1'] = 'Branch'
    sme_sheet['A1'].font = Font(size=10, bold=True)
    sme_sheet['A1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    sme_sheet['B1'] = 'Book Size'
    sme_sheet['B1'].font = Font(size=10, bold=True)
    sme_sheet['B1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    sme_sheet['C1'] = 'VAR7'
    sme_sheet['C1'].font = Font(size=10, bold=True)
    sme_sheet['C1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    sme_sheet['D1'] = 'VAR30'
    sme_sheet['D1'].font = Font(size=10, bold=True)
    sme_sheet['D1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    sme_sheet['E1'] = 'PAR7'
    sme_sheet['E1'].font = Font(size=10, bold=True)
    sme_sheet['E1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    sme_sheet['F1'] = 'PAR30'
    sme_sheet['F1'].font = Font(size=10, bold=True)
    sme_sheet['F1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    sme_sheet['G1'] = 'VAR7 ' + delta_icon + ' FRM Yest'
    sme_sheet['G1'].font = Font(size=10, bold=True)
    sme_sheet['G1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    sme_sheet['H1'] = 'VAR7 ' + delta_icon + ' FRM Beg. of Month'
    sme_sheet['H1'].font = Font(size=10, bold=True)
    sme_sheet['H1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    sme_sheet['I1'] = 'VAR30 ' + delta_icon + ' FRM Yest'
    sme_sheet['I1'].font = Font(size=10, bold=True)
    sme_sheet['I1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    sme_sheet['J1'] = 'VAR30 ' + delta_icon + ' FRM Beg. of Month'
    sme_sheet['J1'].font = Font(size=10, bold=True)
    sme_sheet['J1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    sme_sheet['K1'] = 'No. of accs. in Arrears'
    sme_sheet['K1'].font = Font(size=10, bold=True)
    sme_sheet['K1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    sme_sheet['L1'] = 'Upcoming 3 days'
    sme_sheet['L1'].font = Font(size=10, bold=True)
    sme_sheet['L1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    sme_sheet['M1'] = 'Upcoming 7 days'
    sme_sheet['M1'].font = Font(size=10, bold=True)
    sme_sheet['M1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    sme_sheet['N1'] = 'Upcoming 14 days'
    sme_sheet['N1'].font = Font(size=10, bold=True)
    sme_sheet['N1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    sme_sheet['O1'] = 'Upcoming 21 days'
    sme_sheet['O1'].font = Font(size=10, bold=True)
    sme_sheet['O1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')

def sme_agents_headings(sme_agent_sheet):
    sme_agent_sheet['A1'] = 'Loan Officers'
    sme_agent_sheet['A1'].font = Font(size=10, bold=True)
    sme_agent_sheet['A1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    sme_agent_sheet['B1'] = 'Book Size'
    sme_agent_sheet['B1'].font = Font(size=10, bold=True)
    sme_agent_sheet['B1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    sme_agent_sheet['C1'] = 'VAR7'
    sme_agent_sheet['C1'].font = Font(size=10, bold=True)
    sme_agent_sheet['C1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    sme_agent_sheet['D1'] = 'VAR30'
    sme_agent_sheet['D1'].font = Font(size=10, bold=True)
    sme_agent_sheet['D1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    sme_agent_sheet['E1'] = 'PAR7'
    sme_agent_sheet['E1'].font = Font(size=10, bold=True)
    sme_agent_sheet['E1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    sme_agent_sheet['F1'] = 'PAR30'
    sme_agent_sheet['F1'].font = Font(size=10, bold=True)
    sme_agent_sheet['F1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    sme_agent_sheet['G1'] = 'VAR7 ' + delta_icon + ' FRM Yest'
    sme_agent_sheet['G1'].font = Font(size=10, bold=True)
    sme_agent_sheet['G1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    sme_agent_sheet['H1'] = 'VAR7 ' + delta_icon + ' FRM Beg. of Month'
    sme_agent_sheet['H1'].font = Font(size=10, bold=True)
    sme_agent_sheet['H1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    sme_agent_sheet['I1'] = 'VAR30 ' + delta_icon + ' FRM Yest'
    sme_agent_sheet['I1'].font = Font(size=10, bold=True)
    sme_agent_sheet['I1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    sme_agent_sheet['J1'] = 'VAR30 ' + delta_icon + ' FRM Beg. of Month'
    sme_agent_sheet['J1'].font = Font(size=10, bold=True)
    sme_agent_sheet['J1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    sme_agent_sheet['K1'] = 'No. of accs. in Arrears'
    sme_agent_sheet['K1'].font = Font(size=10, bold=True)
    sme_agent_sheet['K1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    sme_agent_sheet['L1'] = 'Upcoming 3 days'
    sme_agent_sheet['L1'].font = Font(size=10, bold=True)
    sme_agent_sheet['L1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    sme_agent_sheet['M1'] = 'Upcoming 7 days'
    sme_agent_sheet['M1'].font = Font(size=10, bold=True)
    sme_agent_sheet['M1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    sme_agent_sheet['N1'] = 'Upcoming 14 days'
    sme_agent_sheet['N1'].font = Font(size=10, bold=True)
    sme_agent_sheet['N1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    sme_agent_sheet['O1'] = 'Upcoming 21 days'
    sme_agent_sheet['O1'].font = Font(size=10, bold=True)
    sme_agent_sheet['O1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')

def easy_phone_headings(easy_phone_sheet):
    easy_phone_sheet['A1'] = 'Sales Agent'
    easy_phone_sheet['A1'].font = Font(size=10, bold=True)
    easy_phone_sheet['A1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    easy_phone_sheet['B1'] = 'Book Size'
    easy_phone_sheet['B1'].font = Font(size=10, bold=True)
    easy_phone_sheet['B1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    easy_phone_sheet['C1'] = 'VAR3'
    easy_phone_sheet['C1'].font = Font(size=10, bold=True)
    easy_phone_sheet['C1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    easy_phone_sheet['D1'] = 'VAR7'
    easy_phone_sheet['D1'].font = Font(size=10, bold=True)
    easy_phone_sheet['D1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    easy_phone_sheet['E1'] = 'VAR30'
    easy_phone_sheet['E1'].font = Font(size=10, bold=True)
    easy_phone_sheet['E1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    easy_phone_sheet['F1'] = 'PAR3'
    easy_phone_sheet['F1'].font = Font(size=10, bold=True)
    easy_phone_sheet['F1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    easy_phone_sheet['G1'] = 'PAR7'
    easy_phone_sheet['G1'].font = Font(size=10, bold=True)
    easy_phone_sheet['G1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    easy_phone_sheet['H1'] = 'PAR30'
    easy_phone_sheet['H1'].font = Font(size=10, bold=True)
    easy_phone_sheet['H1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    easy_phone_sheet['I1'] = 'VAR7 ' + delta_icon + ' FRM Yest'
    easy_phone_sheet['I1'].font = Font(size=10, bold=True)
    easy_phone_sheet['I1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    easy_phone_sheet['J1'] = 'VAR7 ' + delta_icon + ' FRM Beg. of Month'
    easy_phone_sheet['J1'].font = Font(size=10, bold=True)
    easy_phone_sheet['J1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    easy_phone_sheet['K1'] = 'VAR30 ' + delta_icon + ' FRM Yest'
    easy_phone_sheet['K1'].font = Font(size=10, bold=True)
    easy_phone_sheet['K1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    easy_phone_sheet['L1'] = 'VAR30 ' + delta_icon + ' FRM Beg. of Month'
    easy_phone_sheet['L1'].font = Font(size=10, bold=True)
    easy_phone_sheet['L1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    easy_phone_sheet['M1'] = 'Book size ' + delta_icon + ' FRM Beg. Yest.'
    easy_phone_sheet['M1'].font = Font(size=10, bold=True)
    easy_phone_sheet['M1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    easy_phone_sheet['N1'] = 'Book size ' + delta_icon + ' FRM Beg. of Month'
    easy_phone_sheet['N1'].font = Font(size=10, bold=True)
    easy_phone_sheet['N1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    easy_phone_sheet['O1'] = 'Book size ' + delta_icon + ' FRM Beg. of Year'
    easy_phone_sheet['O1'].font = Font(size=10, bold=True)
    easy_phone_sheet['O1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')

def working_capital_headings(working_capital_sheet):
    working_capital_sheet['A1'] = 'Sales Agent'
    working_capital_sheet['A1'].font = Font(size=10, bold=True)
    working_capital_sheet['A1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    working_capital_sheet['B1'] = 'Book Size'
    working_capital_sheet['B1'].font = Font(size=10, bold=True)
    working_capital_sheet['B1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    working_capital_sheet['C1'] = 'VAR3'
    working_capital_sheet['C1'].font = Font(size=10, bold=True)
    working_capital_sheet['C1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    working_capital_sheet['D1'] = 'VAR7'
    working_capital_sheet['D1'].font = Font(size=10, bold=True)
    working_capital_sheet['D1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    working_capital_sheet['E1'] = 'VAR30'
    working_capital_sheet['E1'].font = Font(size=10, bold=True)
    working_capital_sheet['E1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    working_capital_sheet['F1'] = 'PAR3'
    working_capital_sheet['F1'].font = Font(size=10, bold=True)
    working_capital_sheet['F1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    working_capital_sheet['G1'] = 'PAR7'
    working_capital_sheet['G1'].font = Font(size=10, bold=True)
    working_capital_sheet['G1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    working_capital_sheet['H1'] = 'PAR30'
    working_capital_sheet['H1'].font = Font(size=10, bold=True)
    working_capital_sheet['H1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    working_capital_sheet['I1'] = 'VAR7 ' + delta_icon + ' FRM Yest'
    working_capital_sheet['I1'].font = Font(size=10, bold=True)
    working_capital_sheet['I1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    working_capital_sheet['J1'] = 'VAR7 ' + delta_icon + ' FRM Beg. of Month'
    working_capital_sheet['J1'].font = Font(size=10, bold=True)
    working_capital_sheet['J1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    working_capital_sheet['K1'] = 'VAR30 ' + delta_icon + ' FRM Yest'
    working_capital_sheet['K1'].font = Font(size=10, bold=True)
    working_capital_sheet['K1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    working_capital_sheet['L1'] = 'VAR30 ' + delta_icon + ' FRM Beg. of Month'
    working_capital_sheet['L1'].font = Font(size=10, bold=True)
    working_capital_sheet['L1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    working_capital_sheet['M1'] = 'Book size ' + delta_icon + ' FRM Beg. Yest.'
    working_capital_sheet['M1'].font = Font(size=10, bold=True)
    working_capital_sheet['M1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    working_capital_sheet['N1'] = 'Book size ' + delta_icon + ' FRM Beg. of Month'
    working_capital_sheet['N1'].font = Font(size=10, bold=True)
    working_capital_sheet['N1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    working_capital_sheet['O1'] = 'Book size ' + delta_icon + ' FRM Beg. of Year'
    working_capital_sheet['O1'].font = Font(size=10, bold=True)
    working_capital_sheet['O1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')

def write_entry_summary(summary_sheet, A, B, C, D, E, F, G, H, I, row):
    summary_sheet['A' + str(row)] = A
    summary_sheet['A' + str(row)].font = Font(size=9)
    summary_sheet['A' + str(row)].number_format = '@'
    summary_sheet['B' + str(row)] = B
    summary_sheet['B' + str(row)].font = Font(size=9)
    summary_sheet['B' + str(row)].number_format = '#,##0'
    summary_sheet['C' + str(row)] = C
    summary_sheet['C' + str(row)].font = Font(size=9)
    summary_sheet['C' + str(row)].number_format = '#,##0'
    summary_sheet['D' + str(row)] = D
    summary_sheet['D' + str(row)].font = Font(size=9)
    summary_sheet['D' + str(row)].number_format = '0.00%'
    summary_sheet['E' + str(row)] = E
    summary_sheet['E' + str(row)].font = Font(size=9)
    summary_sheet['E' + str(row)].number_format = '#,##0'
    summary_sheet['F' + str(row)] = F
    summary_sheet['F' + str(row)].font = Font(size=9)
    summary_sheet['F' + str(row)].number_format = '#,##0'
    summary_sheet['G' + str(row)] = G
    summary_sheet['G' + str(row)].font = Font(size=9)
    summary_sheet['G' + str(row)].number_format = '#,##0'
    summary_sheet['H' + str(row)] = H
    summary_sheet['H' + str(row)].font = Font(size=9)
    summary_sheet['H' + str(row)].number_format = '#,##0'
    summary_sheet['I' + str(row)] = I
    summary_sheet['I' + str(row)].font = Font(size=9)
    summary_sheet['I' + str(row)].number_format = '#,##0'

def write_entry_collection(collections_sheet, A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, row):
    collections_sheet['A' + str(row)] = A
    collections_sheet['A' + str(row)].font = Font(size=9)
    collections_sheet['A' + str(row)].number_format = '@'
    collections_sheet['B' + str(row)] = B
    collections_sheet['B' + str(row)].font = Font(size=9)
    collections_sheet['B' + str(row)].number_format = '#,##0'
    collections_sheet['C' + str(row)] = C
    collections_sheet['C' + str(row)].font = Font(size=9)
    collections_sheet['C' + str(row)].number_format = '#,##0'
    collections_sheet['D' + str(row)] = D
    collections_sheet['D' + str(row)].font = Font(size=9)
    collections_sheet['D' + str(row)].number_format = '0.00%'
    collections_sheet['E' + str(row)] = E
    collections_sheet['E' + str(row)].font = Font(size=9)
    collections_sheet['E' + str(row)].number_format = '#,##0'
    collections_sheet['F' + str(row)] = F
    collections_sheet['F' + str(row)].font = Font(size=9)
    collections_sheet['F' + str(row)].number_format = '#,##0'
    collections_sheet['G' + str(row)] = G
    collections_sheet['G' + str(row)].font = Font(size=9)
    collections_sheet['G' + str(row)].number_format = '#,##0'
    collections_sheet['H' + str(row)] = H
    collections_sheet['H' + str(row)].font = Font(size=9)
    collections_sheet['H' + str(row)].number_format = '#,##0'
    collections_sheet['I' + str(row)] = I
    collections_sheet['I' + str(row)].font = Font(size=9)
    collections_sheet['I' + str(row)].number_format = '#,##0'
    collections_sheet['J' + str(row)] = J
    collections_sheet['J' + str(row)].font = Font(size=9)
    collections_sheet['J' + str(row)].number_format = '#,##0'
    collections_sheet['K' + str(row)] = K
    collections_sheet['K' + str(row)].font = Font(size=9)
    collections_sheet['K' + str(row)].number_format = '#,##0'
    collections_sheet['L' + str(row)] = L
    collections_sheet['L' + str(row)].font = Font(size=9)
    collections_sheet['L' + str(row)].number_format = '#,##0'
    collections_sheet['M' + str(row)] = M
    collections_sheet['M' + str(row)].font = Font(size=9)
    collections_sheet['M' + str(row)].number_format = '#,##0'
    collections_sheet['N' + str(row)] = N
    collections_sheet['N' + str(row)].font = Font(size=9)
    collections_sheet['N' + str(row)].number_format = '#,##0'
    collections_sheet['O' + str(row)] = O
    collections_sheet['O' + str(row)].font = Font(size=9)
    collections_sheet['O' + str(row)].number_format = '#,##0'

def write_entry_branch_summary(branch_summary_sheet, A, B, C, D, E, F, G, H, I, J, K, L, M, row):
    branch_summary_sheet['A' + str(row)] = A
    branch_summary_sheet['A' + str(row)].font = Font(size=9)
    branch_summary_sheet['A' + str(row)].number_format = '@'
    branch_summary_sheet['B' + str(row)] = B
    branch_summary_sheet['B' + str(row)].font = Font(size=9)
    branch_summary_sheet['B' + str(row)].number_format = '#,##0'
    branch_summary_sheet['C' + str(row)] = C
    branch_summary_sheet['C' + str(row)].font = Font(size=9)
    branch_summary_sheet['C' + str(row)].number_format = '#,##0'
    branch_summary_sheet['D' + str(row)] = D
    branch_summary_sheet['D' + str(row)].font = Font(size=9)
    branch_summary_sheet['D' + str(row)].number_format = '0.00%'
    branch_summary_sheet['E' + str(row)] = E
    branch_summary_sheet['E' + str(row)].font = Font(size=9)
    branch_summary_sheet['E' + str(row)].number_format = '0.00%'
    branch_summary_sheet['F' + str(row)] = F
    branch_summary_sheet['F' + str(row)].font = Font(size=9)
    branch_summary_sheet['F' + str(row)].number_format = '#,##0'
    branch_summary_sheet['G' + str(row)] = G
    branch_summary_sheet['G' + str(row)].font = Font(size=9)
    branch_summary_sheet['G' + str(row)].number_format = '#,##0'
    branch_summary_sheet['H' + str(row)] = H
    branch_summary_sheet['H' + str(row)].font = Font(size=9)
    branch_summary_sheet['H' + str(row)].number_format = '#,##0'
    branch_summary_sheet['I' + str(row)] = I
    branch_summary_sheet['I' + str(row)].font = Font(size=9)
    branch_summary_sheet['I' + str(row)].number_format = '#,##0'
    branch_summary_sheet['J' + str(row)] = J
    branch_summary_sheet['J' + str(row)].font = Font(size=9)
    branch_summary_sheet['J' + str(row)].number_format = '#,##0'
    branch_summary_sheet['K' + str(row)] = K
    branch_summary_sheet['K' + str(row)].font = Font(size=9)
    branch_summary_sheet['K' + str(row)].number_format = '#,##0'
    branch_summary_sheet['L' + str(row)] = L
    branch_summary_sheet['L' + str(row)].font = Font(size=9)
    branch_summary_sheet['L' + str(row)].number_format = '#,##0'
    branch_summary_sheet['M' + str(row)] = M
    branch_summary_sheet['M' + str(row)].font = Font(size=9)
    branch_summary_sheet['M' + str(row)].number_format = '#,##0'

def write_entry_lbf(lbf_sheet, A, B, C, D, E, F, G, H, I, J, K, L, M, N, row):
    lbf_sheet['A' + str(row)] = A
    lbf_sheet['A' + str(row)].font = Font(size=9)
    lbf_sheet['A' + str(row)].number_format = '@'
    lbf_sheet['B' + str(row)] = B
    lbf_sheet['B' + str(row)].font = Font(size=9)
    lbf_sheet['B' + str(row)].number_format = '#,##0'
    lbf_sheet['C' + str(row)] = C
    lbf_sheet['C' + str(row)].font = Font(size=9)
    lbf_sheet['C' + str(row)].number_format = '#,##0'
    lbf_sheet['D' + str(row)] = D
    lbf_sheet['D' + str(row)].font = Font(size=9)
    lbf_sheet['D' + str(row)].number_format = '0.00%'
    lbf_sheet['E' + str(row)] = E
    lbf_sheet['E' + str(row)].font = Font(size=9)
    lbf_sheet['E' + str(row)].number_format = '#,##0'
    lbf_sheet['F' + str(row)] = F
    lbf_sheet['F' + str(row)].font = Font(size=9)
    lbf_sheet['F' + str(row)].number_format = '0.00%'
    lbf_sheet['G' + str(row)] = G
    lbf_sheet['G' + str(row)].font = Font(size=9)
    lbf_sheet['G' + str(row)].number_format = '#,##0'
    lbf_sheet['H' + str(row)] = H
    lbf_sheet['H' + str(row)].font = Font(size=9)
    lbf_sheet['H' + str(row)].number_format = '#,##0'
    lbf_sheet['I' + str(row)] = I
    lbf_sheet['I' + str(row)].font = Font(size=9)
    lbf_sheet['I' + str(row)].number_format = '#,##0'
    lbf_sheet['J' + str(row)] = J
    lbf_sheet['J' + str(row)].font = Font(size=9)
    lbf_sheet['J' + str(row)].number_format = '#,##0'
    lbf_sheet['K' + str(row)] = K
    lbf_sheet['K' + str(row)].font = Font(size=9)
    lbf_sheet['K' + str(row)].number_format = '#,##0'
    lbf_sheet['L' + str(row)] = L
    lbf_sheet['L' + str(row)].font = Font(size=9)
    lbf_sheet['L' + str(row)].number_format = '#,##0'
    lbf_sheet['M' + str(row)] = M
    lbf_sheet['M' + str(row)].font = Font(size=9)
    lbf_sheet['M' + str(row)].number_format = '#,##0'
    lbf_sheet['N' + str(row)] = N
    lbf_sheet['N' + str(row)].font = Font(size=9)
    lbf_sheet['N' + str(row)].number_format = '#,##0'

def write_entry_lbf_branch(lbf_sheet, A, B, C, D, E, F, G, H, I, J, K, L, M, N, row):
    lbf_sheet['A' + str(row)] = A
    lbf_sheet['A' + str(row)].font = Font(size=9)
    lbf_sheet['A' + str(row)].number_format = '@'
    lbf_sheet['B' + str(row)] = B
    lbf_sheet['B' + str(row)].font = Font(size=9)
    lbf_sheet['B' + str(row)].number_format = '#,##0'
    lbf_sheet['C' + str(row)] = C
    lbf_sheet['C' + str(row)].font = Font(size=9)
    lbf_sheet['C' + str(row)].number_format = '#,##0'
    lbf_sheet['D' + str(row)] = D
    lbf_sheet['D' + str(row)].font = Font(size=9)
    lbf_sheet['D' + str(row)].number_format = '0.00%'
    lbf_sheet['E' + str(row)] = E
    lbf_sheet['E' + str(row)].font = Font(size=9)
    lbf_sheet['E' + str(row)].number_format = '#,##0'
    lbf_sheet['F' + str(row)] = F
    lbf_sheet['F' + str(row)].font = Font(size=9)
    lbf_sheet['F' + str(row)].number_format = '0.00%'
    lbf_sheet['G' + str(row)] = G
    lbf_sheet['G' + str(row)].font = Font(size=9)
    lbf_sheet['G' + str(row)].number_format = '#,##0'
    lbf_sheet['H' + str(row)] = H
    lbf_sheet['H' + str(row)].font = Font(size=9)
    lbf_sheet['H' + str(row)].number_format = '#,##0'
    lbf_sheet['I' + str(row)] = I
    lbf_sheet['I' + str(row)].font = Font(size=9)
    lbf_sheet['I' + str(row)].number_format = '#,##0'
    lbf_sheet['J' + str(row)] = J
    lbf_sheet['J' + str(row)].font = Font(size=9)
    lbf_sheet['J' + str(row)].number_format = '#,##0'
    lbf_sheet['K' + str(row)] = K
    lbf_sheet['K' + str(row)].font = Font(size=9)
    lbf_sheet['K' + str(row)].number_format = '#,##0'
    lbf_sheet['L' + str(row)] = L
    lbf_sheet['L' + str(row)].font = Font(size=9)
    lbf_sheet['L' + str(row)].number_format = '#,##0'
    lbf_sheet['M' + str(row)] = M
    lbf_sheet['M' + str(row)].font = Font(size=9)
    lbf_sheet['M' + str(row)].number_format = '#,##0'
    lbf_sheet['N' + str(row)] = N
    lbf_sheet['N' + str(row)].font = Font(size=9)
    lbf_sheet['N' + str(row)].number_format = '#,##0'

def write_entry_private_sector_client(private_sector_sheet, A, B, C, D, E, F, G, H, row):
    private_sector_sheet['A' + str(row)] = A
    private_sector_sheet['A' + str(row)].font = Font(size=9)
    private_sector_sheet['A' + str(row)].number_format = '@'
    private_sector_sheet['B' + str(row)] = B
    private_sector_sheet['B' + str(row)].font = Font(size=9)
    private_sector_sheet['B' + str(row)].number_format = '#,##0'
    private_sector_sheet['C' + str(row)] = C
    private_sector_sheet['C' + str(row)].font = Font(size=9)
    private_sector_sheet['C' + str(row)].number_format = '#,##0'
    private_sector_sheet['D' + str(row)] = D
    private_sector_sheet['D' + str(row)].font = Font(size=9)
    private_sector_sheet['D' + str(row)].number_format = '0.00%'
    private_sector_sheet['E' + str(row)] = E
    private_sector_sheet['E' + str(row)].font = Font(size=9)
    private_sector_sheet['E' + str(row)].number_format = '#,##0'
    private_sector_sheet['F' + str(row)] = F
    private_sector_sheet['F' + str(row)].font = Font(size=9)
    private_sector_sheet['F' + str(row)].number_format = '#,##0'
    private_sector_sheet['G' + str(row)] = G
    private_sector_sheet['G' + str(row)].font = Font(size=9)
    private_sector_sheet['G' + str(row)].number_format = '#,##0'
    private_sector_sheet['H' + str(row)] = H
    private_sector_sheet['H' + str(row)].font = Font(size=9)
    private_sector_sheet['H' + str(row)].number_format = '#,##0'

def write_entry_sme(sme_sheet, A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, row):
    sme_sheet['A' + str(row)] = A
    sme_sheet['A' + str(row)].font = Font(size=9)
    sme_sheet['A' + str(row)].number_format = '@'
    sme_sheet['B' + str(row)] = B
    sme_sheet['B' + str(row)].font = Font(size=9)
    sme_sheet['B' + str(row)].number_format = '#,##0'
    sme_sheet['C' + str(row)] = C
    sme_sheet['C' + str(row)].font = Font(size=9)
    sme_sheet['C' + str(row)].number_format = '#,##0'
    sme_sheet['D' + str(row)] = D
    sme_sheet['D' + str(row)].font = Font(size=9)
    sme_sheet['D' + str(row)].number_format = '#,##0'
    sme_sheet['E' + str(row)] = E
    sme_sheet['E' + str(row)].font = Font(size=9)
    sme_sheet['E' + str(row)].number_format = '0.00%'
    sme_sheet['F' + str(row)] = F
    sme_sheet['F' + str(row)].font = Font(size=9)
    sme_sheet['F' + str(row)].number_format = '0.00%'
    sme_sheet['G' + str(row)] = G
    sme_sheet['G' + str(row)].font = Font(size=9)
    sme_sheet['G' + str(row)].number_format = '#,##0'
    sme_sheet['H' + str(row)] = H
    sme_sheet['H' + str(row)].font = Font(size=9)
    sme_sheet['H' + str(row)].number_format = '#,##0'
    sme_sheet['I' + str(row)] = I
    sme_sheet['I' + str(row)].font = Font(size=9)
    sme_sheet['I' + str(row)].number_format = '#,##0'
    sme_sheet['J' + str(row)] = J
    sme_sheet['J' + str(row)].font = Font(size=9)
    sme_sheet['J' + str(row)].number_format = '#,##0'
    sme_sheet['K' + str(row)] = K
    sme_sheet['K' + str(row)].font = Font(size=9)
    sme_sheet['K' + str(row)].number_format = '#,##0'
    sme_sheet['L' + str(row)] = L
    sme_sheet['L' + str(row)].font = Font(size=9)
    sme_sheet['L' + str(row)].number_format = '#,##0'
    sme_sheet['M' + str(row)] = M
    sme_sheet['M' + str(row)].font = Font(size=9)
    sme_sheet['M' + str(row)].number_format = '#,##0'
    sme_sheet['N' + str(row)] = N
    sme_sheet['N' + str(row)].font = Font(size=9)
    sme_sheet['N' + str(row)].number_format = '#,##0'
    sme_sheet['O' + str(row)] = O
    sme_sheet['O' + str(row)].font = Font(size=9)
    sme_sheet['O' + str(row)].number_format = '#,##0'

def write_entry_sme_agent(sme_agent_sheet, A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, row):
    sme_agent_sheet['A' + str(row)] = A
    sme_agent_sheet['A' + str(row)].font = Font(size=9)
    sme_agent_sheet['A' + str(row)].number_format = '@'
    sme_agent_sheet['B' + str(row)] = B
    sme_agent_sheet['B' + str(row)].font = Font(size=9)
    sme_agent_sheet['B' + str(row)].number_format = '#,##0'
    sme_agent_sheet['C' + str(row)] = C
    sme_agent_sheet['C' + str(row)].font = Font(size=9)
    sme_agent_sheet['C' + str(row)].number_format = '#,##0'
    sme_agent_sheet['D' + str(row)] = D
    sme_agent_sheet['D' + str(row)].font = Font(size=9)
    sme_agent_sheet['D' + str(row)].number_format = '#,##0'
    sme_agent_sheet['E' + str(row)] = E
    sme_agent_sheet['E' + str(row)].font = Font(size=9)
    sme_agent_sheet['E' + str(row)].number_format = '0.00%'
    sme_agent_sheet['F' + str(row)] = F
    sme_agent_sheet['F' + str(row)].font = Font(size=9)
    sme_agent_sheet['F' + str(row)].number_format = '0.00%'
    sme_agent_sheet['G' + str(row)] = G
    sme_agent_sheet['G' + str(row)].font = Font(size=9)
    sme_agent_sheet['G' + str(row)].number_format = '#,##0'
    sme_agent_sheet['H' + str(row)] = H
    sme_agent_sheet['H' + str(row)].font = Font(size=9)
    sme_agent_sheet['H' + str(row)].number_format = '#,##0'
    sme_agent_sheet['I' + str(row)] = I
    sme_agent_sheet['I' + str(row)].font = Font(size=9)
    sme_agent_sheet['I' + str(row)].number_format = '#,##0'
    sme_agent_sheet['J' + str(row)] = J
    sme_agent_sheet['J' + str(row)].font = Font(size=9)
    sme_agent_sheet['J' + str(row)].number_format = '#,##0'
    sme_agent_sheet['K' + str(row)] = K
    sme_agent_sheet['K' + str(row)].font = Font(size=9)
    sme_agent_sheet['K' + str(row)].number_format = '#,##0'
    sme_agent_sheet['L' + str(row)] = L
    sme_agent_sheet['L' + str(row)].font = Font(size=9)
    sme_agent_sheet['L' + str(row)].number_format = '#,##0'
    sme_agent_sheet['M' + str(row)] = M
    sme_agent_sheet['M' + str(row)].font = Font(size=9)
    sme_agent_sheet['M' + str(row)].number_format = '#,##0'
    sme_agent_sheet['N' + str(row)] = N
    sme_agent_sheet['N' + str(row)].font = Font(size=9)
    sme_agent_sheet['N' + str(row)].number_format = '#,##0'
    sme_agent_sheet['O' + str(row)] = O
    sme_agent_sheet['O' + str(row)].font = Font(size=9)
    sme_agent_sheet['O' + str(row)].number_format = '#,##0'

def write_entry_easy_phone(easy_phone_sheet, A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, row):
    easy_phone_sheet['A' + str(row)] = A
    easy_phone_sheet['A' + str(row)].font = Font(size=9)
    easy_phone_sheet['A' + str(row)].number_format = '@'
    easy_phone_sheet['B' + str(row)] = B
    easy_phone_sheet['B' + str(row)].font = Font(size=9)
    easy_phone_sheet['B' + str(row)].number_format = '#,##0'
    easy_phone_sheet['C' + str(row)] = C
    easy_phone_sheet['C' + str(row)].font = Font(size=9)
    easy_phone_sheet['C' + str(row)].number_format = '#,##0'
    easy_phone_sheet['D' + str(row)] = D
    easy_phone_sheet['D' + str(row)].font = Font(size=9)
    easy_phone_sheet['D' + str(row)].number_format = '#,##0'
    easy_phone_sheet['E' + str(row)] = E
    easy_phone_sheet['E' + str(row)].font = Font(size=9)
    easy_phone_sheet['E' + str(row)].number_format = '#,##0'
    easy_phone_sheet['F' + str(row)] = F
    easy_phone_sheet['F' + str(row)].font = Font(size=9)
    easy_phone_sheet['F' + str(row)].number_format = '0.00%'
    easy_phone_sheet['G' + str(row)] = G
    easy_phone_sheet['G' + str(row)].font = Font(size=9)
    easy_phone_sheet['G' + str(row)].number_format = '0.00%'
    easy_phone_sheet['H' + str(row)] = H
    easy_phone_sheet['H' + str(row)].font = Font(size=9)
    easy_phone_sheet['H' + str(row)].number_format = '0.00%'
    easy_phone_sheet['I' + str(row)] = I
    easy_phone_sheet['I' + str(row)].font = Font(size=9)
    easy_phone_sheet['I' + str(row)].number_format = '#,##0'
    easy_phone_sheet['J' + str(row)] = J
    easy_phone_sheet['J' + str(row)].font = Font(size=9)
    easy_phone_sheet['J' + str(row)].number_format = '#,##0'
    easy_phone_sheet['K' + str(row)] = K
    easy_phone_sheet['K' + str(row)].font = Font(size=9)
    easy_phone_sheet['K' + str(row)].number_format = '#,##0'
    easy_phone_sheet['L' + str(row)] = L
    easy_phone_sheet['L' + str(row)].font = Font(size=9)
    easy_phone_sheet['L' + str(row)].number_format = '#,##0'
    easy_phone_sheet['M' + str(row)] = M
    easy_phone_sheet['M' + str(row)].font = Font(size=9)
    easy_phone_sheet['M' + str(row)].number_format = '#,##0'
    easy_phone_sheet['N' + str(row)] = N
    easy_phone_sheet['N' + str(row)].font = Font(size=9)
    easy_phone_sheet['N' + str(row)].number_format = '#,##0'
    easy_phone_sheet['O' + str(row)] = O
    easy_phone_sheet['O' + str(row)].font = Font(size=9)
    easy_phone_sheet['O' + str(row)].number_format = '#,##0'

def write_entry_working_capital(working_capital_sheet, A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, row):
    working_capital_sheet['A' + str(row)] = A
    working_capital_sheet['A' + str(row)].font = Font(size=9)
    working_capital_sheet['A' + str(row)].number_format = '@'
    working_capital_sheet['B' + str(row)] = B
    working_capital_sheet['B' + str(row)].font = Font(size=9)
    working_capital_sheet['B' + str(row)].number_format = '#,##0'
    working_capital_sheet['C' + str(row)] = C
    working_capital_sheet['C' + str(row)].font = Font(size=9)
    working_capital_sheet['C' + str(row)].number_format = '#,##0'
    working_capital_sheet['D' + str(row)] = D
    working_capital_sheet['D' + str(row)].font = Font(size=9)
    working_capital_sheet['D' + str(row)].number_format = '#,##0'
    working_capital_sheet['E' + str(row)] = E
    working_capital_sheet['E' + str(row)].font = Font(size=9)
    working_capital_sheet['E' + str(row)].number_format = '#,##0'
    working_capital_sheet['F' + str(row)] = F
    working_capital_sheet['F' + str(row)].font = Font(size=9)
    working_capital_sheet['F' + str(row)].number_format = '0.00%'
    working_capital_sheet['G' + str(row)] = G
    working_capital_sheet['G' + str(row)].font = Font(size=9)
    working_capital_sheet['G' + str(row)].number_format = '0.00%'
    working_capital_sheet['H' + str(row)] = H
    working_capital_sheet['H' + str(row)].font = Font(size=9)
    working_capital_sheet['H' + str(row)].number_format = '0.00%'
    working_capital_sheet['I' + str(row)] = I
    working_capital_sheet['I' + str(row)].font = Font(size=9)
    working_capital_sheet['I' + str(row)].number_format = '#,##0'
    working_capital_sheet['J' + str(row)] = J
    working_capital_sheet['J' + str(row)].font = Font(size=9)
    working_capital_sheet['J' + str(row)].number_format = '#,##0'
    working_capital_sheet['K' + str(row)] = K
    working_capital_sheet['K' + str(row)].font = Font(size=9)
    working_capital_sheet['K' + str(row)].number_format = '#,##0'
    working_capital_sheet['L' + str(row)] = L
    working_capital_sheet['L' + str(row)].font = Font(size=9)
    working_capital_sheet['L' + str(row)].number_format = '#,##0'
    working_capital_sheet['M' + str(row)] = M
    working_capital_sheet['M' + str(row)].font = Font(size=9)
    working_capital_sheet['M' + str(row)].number_format = '#,##0'
    working_capital_sheet['N' + str(row)] = N
    working_capital_sheet['N' + str(row)].font = Font(size=9)
    working_capital_sheet['N' + str(row)].number_format = '#,##0'
    working_capital_sheet['O' + str(row)] = O
    working_capital_sheet['O' + str(row)].font = Font(size=9)
    working_capital_sheet['O' + str(row)].number_format = '#,##0'

def product_book_size(product, loans_df):
    loans_filter_df = loans_df[loans_df['Products'] == product]
    total_portfolio = loans_filter_df['Principal Balance'].sum()
    return total_portfolio

def product_par(product, days, loans_df):
    product_filter = loans_df[loans_df['Products'] == product]
    value_in_arr = product_filter[product_filter['Days In Arrears'] > days]['Principal Balance'].sum()
    portfolio = product_filter['Principal Balance'].sum()
    par = value_in_arr / portfolio if portfolio > 0 else 0
    return (value_in_arr, par)

def collection_book_size(collection_officer, loans_df):
    loans_filter_df = loans_df[loans_df['Collection_Officer'] == collection_officer]
    total_portfolio = loans_filter_df['Principal Balance'].sum()
    return total_portfolio

def collection_par(collection_officer, days, loans_df):
    agent_filter_df = loans_df[loans_df['Collection_Officer'] == collection_officer]
    value_in_arr = agent_filter_df[agent_filter_df['Days In Arrears'] > days]['Principal Balance'].sum()
    portfolio = agent_filter_df['Principal Balance'].sum()
    par = value_in_arr / portfolio if portfolio > 0 else 0
    return (value_in_arr, par)

def collection_by_product_breakdown(collection_officer, loans_df):
    agent_filter_df = loans_df[loans_df['Collection_Officer'] == collection_officer]
    performing_loans = agent_filter_df[agent_filter_df['Collections Status'] == 'Performing/ Normal']['Principal Balance'].sum()
    partial1_loans = agent_filter_df[agent_filter_df['Collections Status'] == 'Partial Installment 1']['Principal Balance'].sum()
    partial2_loans = agent_filter_df[agent_filter_df['Collections Status'] == 'Partial Installment 2']['Principal Balance'].sum()
    npl_vote_loans = agent_filter_df[agent_filter_df['Collections Status'] == 'NPL-Vote']['Principal Balance'].sum()
    npl_vote_client = agent_filter_df[agent_filter_df['Collections Status'] == 'NPL-Client']['Principal Balance'].sum()
    zero_receipts = agent_filter_df[agent_filter_df['Collections Status'] == 'Zero Receipted']['Principal Balance'].sum()
    return (performing_loans, partial1_loans, partial2_loans, npl_vote_loans, npl_vote_client, zero_receipts)

def cs_branch_summ_book_size(branch, loans_df):
    loans_filter_df = loans_df[loans_df['Branch'] == branch]
    total_portfolio = loans_filter_df['Principal Balance'].sum()
    return total_portfolio

def cs_branch_summ_par(branch, days, loans_df):
    branch_filter_df = loans_df[loans_df['Branch'] == branch]
    value_in_arr = branch_filter_df[branch_filter_df['Days In Arrears'] > days]['Principal Balance'].sum()
    portfolio = branch_filter_df['Principal Balance'].sum()
    par = value_in_arr / portfolio if portfolio > 0 else 0
    return (value_in_arr, par)

def cs_branch_summ_bp_par(branch, loans_df):
    branch_filter_df = loans_df[loans_df['Branch'] == branch]
    total_book = branch_filter_df['Principal Balance'].sum()
    other_loans = branch_filter_df[(branch_filter_df['Collections Status'] != 'Performing/ Normal') & branch_filter_df['Collections Status'].notna()]['Principal Balance'].sum()
    bp_par = other_loans / total_book if total_book > 0 else 0
    return bp_par

def cs_branch_by_product_breakdown(branch, loans_df):
    branch_filter_df = loans_df[loans_df['Branch'] == branch]
    performing_loans = branch_filter_df[branch_filter_df['Collections Status'] == 'Performing/ Normal']['Principal Balance'].sum()
    partial1_loans = branch_filter_df[branch_filter_df['Collections Status'] == 'Partial Installment 1']['Principal Balance'].sum()
    partial2_loans = branch_filter_df[branch_filter_df['Collections Status'] == 'Partial Installment 2']['Principal Balance'].sum()
    npl_vote_loans = branch_filter_df[branch_filter_df['Collections Status'] == 'NPL-Vote']['Principal Balance'].sum()
    npl_vote_client = branch_filter_df[branch_filter_df['Collections Status'] == 'NPL-Client']['Principal Balance'].sum()
    zero_receipts = branch_filter_df[branch_filter_df['Collections Status'] == 'Zero Receipted']['Principal Balance'].sum()
    return (performing_loans, partial1_loans, partial2_loans, npl_vote_loans, npl_vote_client, zero_receipts)

def lbf_book_size(ro, loans_df):
    loans_filter_df = loans_df[loans_df['Relationship/Recoveries Officers'] == ro]
    total_portfolio = loans_filter_df['Principal Balance'].sum()
    return total_portfolio

def lbf_par(ro, days, loans_df):
    agent_filter_df = loans_df[loans_df['Relationship/Recoveries Officers'] == ro]
    value_in_arr = agent_filter_df[agent_filter_df['Days In Arrears'] > days]['Principal Balance'].sum()
    portfolio = agent_filter_df['Principal Balance'].sum()
    par = value_in_arr / portfolio if portfolio > 0 else 0
    return (value_in_arr, par)

def lbf_overall_substatus(ro, loans_df):
    agent_filter_df = loans_df[loans_df['Relationship/Recoveries Officers'] == ro]
    repossession = agent_filter_df[agent_filter_df['LBF Sub Status'] == 'Repossession']['Principal Balance'].sum()
    yard = agent_filter_df[agent_filter_df['LBF Sub Status'] == 'Yard']['Principal Balance'].sum()
    sold = agent_filter_df[agent_filter_df['LBF Sub Status'] == 'Sold']['Principal Balance'].sum()
    return (repossession, yard, sold)

def lbf_branch_book_size(branch, loans_df):
    loans_filter_df = loans_df[loans_df['Branch'] == branch]
    total_portfolio = loans_filter_df['Principal Balance'].sum()
    return total_portfolio

def lbf_branch_par(branch, days, loans_df):
    agent_filter_df = loans_df[loans_df['Branch'] == branch]
    value_in_arr = agent_filter_df[agent_filter_df['Days In Arrears'] > days]['Principal Balance'].sum()
    portfolio = agent_filter_df['Principal Balance'].sum()
    par = value_in_arr / portfolio if portfolio > 0 else 0
    return (value_in_arr, par)

def lbf_branch_overall_substatus(branch, loans_df):
    agent_filter_df = loans_df[loans_df['Branch'] == branch]
    repossession = agent_filter_df[agent_filter_df['LBF Sub Status'] == 'Repossession']['Principal Balance'].sum()
    yard = agent_filter_df[agent_filter_df['LBF Sub Status'] == 'Yard']['Principal Balance'].sum()
    sold = agent_filter_df[agent_filter_df['LBF Sub Status'] == 'Sold']['Principal Balance'].sum()
    return (repossession, yard, sold)

def employer_branch_book_size(employer, loans_df):
    loans_filter_df = loans_df[(loans_df['Products'] == 'PSL') & (loans_df['Employer_Name (Client)'] == employer)]
    total_portfolio = loans_filter_df['Principal Balance'].sum()
    return total_portfolio

def employer_branch_par(employer, days, loans_df):
    agent_filter_df = loans_df[(loans_df['Products'] == 'PSL') & (loans_df['Employer_Name (Client)'] == employer)]
    value_in_arr = agent_filter_df[agent_filter_df['Days In Arrears'] > days]['Principal Balance'].sum()
    portfolio = agent_filter_df['Principal Balance'].sum()
    par = value_in_arr / portfolio if portfolio > 0 else 0
    return (value_in_arr, par)

def sme_branch_summ_book_size(branch, loans_df):
    branch_filter_df = loans_df[(loans_df['Branch'] == branch) & loans_df['Products'].isin(SME_PRODUCTS)]
    total_portfolio = branch_filter_df['Principal Balance'].sum()
    return total_portfolio

def sme_branch_summ_par(branch, days, loans_df):
    branch_filter_df = loans_df[(loans_df['Branch'] == branch) & loans_df['Products'].isin(SME_PRODUCTS)]
    value_in_arr = branch_filter_df[branch_filter_df['Days In Arrears'] > days]['Principal Balance'].sum()
    portfolio = branch_filter_df['Principal Balance'].sum()
    par = value_in_arr / portfolio if portfolio > 0 else 0
    return (value_in_arr, par)

def sme_branch_par_mv(branch, loans_df):
    branch_filter_df = loans_df[loans_df['Branch'] == branch]
    value_in_arr = branch_filter_df[branch_filter_df['Days In Arrears'] > 30]['Principal Balance'].sum()
    principal_bal = branch_filter_df['Principal Balance'].sum()
    mvmt = value_in_arr - 0.03 * principal_bal
    if mvmt > 0:
        return mvmt
    return None

def sme_accs_in_arrears(branch, loans_df):
    branch_filter_df = loans_df[(loans_df['Branch'] == branch) & loans_df['Products'].isin(SME_PRODUCTS)]
    accs_in_arrs = branch_filter_df[branch_filter_df['Account State'] == 'In Arrears'].shape[0]
    return accs_in_arrs

def sme_agent_summ_book_size(agent, loans_df):
    branch_filter_df = loans_df[(loans_df['Sales Representative (Loan)'] == agent) & loans_df['Products'].isin(SME_PRODUCTS)]
    total_portfolio = branch_filter_df['Principal Balance'].sum()
    return total_portfolio

def sme_agent_summ_par(agent, days, loans_df):
    branch_filter_df = loans_df[(loans_df['Sales Representative (Loan)'] == agent) & loans_df['Products'].isin(SME_PRODUCTS)]
    value_in_arr = branch_filter_df[branch_filter_df['Days In Arrears'] > days]['Principal Balance'].sum()
    portfolio = branch_filter_df['Principal Balance'].sum()
    par = value_in_arr / portfolio if portfolio > 0 else 0
    return (value_in_arr, par)

def sme_agent_par_mv(agent, loans_df):
    branch_filter_df = loans_df[loans_df['Sales Representative (Loan)'] == agent]
    value_in_arr = branch_filter_df[branch_filter_df['Days In Arrears'] > 30]['Principal Balance'].sum()
    principal_bal = branch_filter_df['Principal Balance'].sum()
    mvmt = value_in_arr - 0.03 * principal_bal
    if mvmt > 0:
        return mvmt
    return None

def sme_agent_accs_in_arrears(agent, loans_df):
    branch_filter_df = loans_df[(loans_df['Sales Representative (Loan)'] == agent) & loans_df['Products'].isin(SME_PRODUCTS)]
    accs_in_arrs = branch_filter_df[branch_filter_df['Account State'] == 'In Arrears'].shape[0]
    return accs_in_arrs

def agent_par(agent, days, loans_df):
    agent_filter_df = loans_df[loans_df['Sales Representative (Loan)'] == agent]
    value_in_arr = agent_filter_df[agent_filter_df['Days In Arrears'] > days]['Principal Balance'].sum()
    portfolio = agent_filter_df['Principal Balance'].sum()
    par = value_in_arr / portfolio if portfolio > 0 else 0
    return par

def agent_var(agent, days, loans_df):
    agent_filter_df = loans_df[loans_df['Sales Representative (Loan)'] == agent]
    value_in_arr = agent_filter_df[agent_filter_df['Days In Arrears'] > days]['Principal Balance'].sum()
    portfolio = agent_filter_df['Principal Balance'].sum()
    return portfolio

def easy_phone_agent_var(agent, days, loans_df):
    agent_filter_df = loans_df[loans_df['Sales Representative (Loan)'] == agent]
    value_in_arr = agent_filter_df[agent_filter_df['Days In Arrears'] > days]['Principal Balance'].sum()
    portfolio = agent_filter_df['Principal Balance'].sum()
    return value_in_arr

def product_generate_entry(loans_df, start_of_year_df, start_of_month_df, yesterdays_all_df, product):
    A = product
    B = product_book_size(product, loans_df)
    C, D = product_par(product, 30, loans_df)
    E = C - product_par(product, 30, yesterdays_all_df)[0]
    F = C - product_par(product, 30, start_of_month_df)[0]
    G = C - product_par(product, 30, start_of_year_df)[0]
    H = B - product_book_size(product, start_of_month_df)
    I = B - product_book_size(product, start_of_year_df)
    return (A, B, C, D, E, F, G, H, I)

def collections_generate_entry(loans_df, start_of_year_df, start_of_month_df, yesterdays_all_df, collection_officer):
    A = collection_officer
    B = collection_book_size(collection_officer, loans_df)
    C, D = collection_par(collection_officer, 30, loans_df)
    E, F, G, H, I, J = collection_by_product_breakdown(collection_officer, loans_df)
    K = C - collection_par(collection_officer, 30, yesterdays_all_df)[0]
    L = C - collection_par(collection_officer, 30, start_of_month_df)[0]
    M = C - collection_par(collection_officer, 30, start_of_year_df)[0]
    N = B - collection_book_size(collection_officer, start_of_month_df)
    O = B - collection_book_size(collection_officer, start_of_year_df)
    return (A, B, C, D, E, F, G, H, I, J, K, L, M, N, O)

def cs_branch_summary_generate_entry(loans_df, start_of_year_df, start_of_month_df, yesterdays_all_df, branch):
    loans_df_filter = loans_df[loans_df['Products'] == 'CSL']
    A = branch
    B = cs_branch_summ_book_size(branch, loans_df_filter)
    C, D = cs_branch_summ_par(branch, 30, loans_df_filter)
    E = cs_branch_summ_bp_par(branch, loans_df_filter)
    F, G, H, I, J, K = cs_branch_by_product_breakdown(branch, loans_df_filter)
    L = C - cs_branch_summ_par(branch, 30, yesterdays_all_df)[0]
    M = C - cs_branch_summ_par(branch, 30, start_of_month_df)[0]
    return (A, B, C, D, E, F, G, H, I, J, K, L, M)

def lbf_generate_entry(loans_df, start_of_year_df, start_of_month_df, yesterdays_all_df, ro):
    loans_df_filter = loans_df[loans_df['Products'] == 'LBF']
    A = ro
    B = lbf_book_size(ro, loans_df_filter)
    C, D = lbf_par(ro, 7, loans_df_filter)
    E, F = lbf_par(ro, 30, loans_df_filter)
    G, H, I = lbf_overall_substatus(ro, loans_df_filter)
    J = E - lbf_par(ro, 30, yesterdays_all_df)[0]
    K = E - lbf_par(ro, 30, start_of_month_df)[0]
    L = E - lbf_par(ro, 30, start_of_year_df)[0]
    M = B - lbf_book_size(ro, start_of_month_df)
    N = B - lbf_book_size(ro, start_of_year_df)
    return (A, B, C, D, E, F, G, H, I, J, K, L, M, N)

def lbf_branch_generate_entry(loans_df, start_of_year_df, start_of_month_df, yesterdays_all_df, branch):
    loans_df_filter = loans_df[loans_df['Products'] == 'LBF']
    A = branch
    B = lbf_branch_book_size(branch, loans_df_filter)
    C, D = lbf_branch_par(branch, 7, loans_df_filter)
    E, F = lbf_branch_par(branch, 30, loans_df_filter)
    G, H, I = lbf_branch_overall_substatus(branch, loans_df_filter)
    J = E - lbf_branch_par(branch, 30, yesterdays_all_df)[0]
    K = E - lbf_branch_par(branch, 30, start_of_month_df)[0]
    L = E - lbf_branch_par(branch, 30, start_of_year_df)[0]
    M = B - lbf_branch_book_size(branch, start_of_month_df)
    N = B - lbf_branch_book_size(branch, start_of_year_df)
    return (A, B, C, D, E, F, G, H, I, J, K, L, M, N)

def private_sector_generate_entry(loans_df, start_of_year_df, start_of_month_df, yesterdays_all_df, employer):
    loans_df_filter = loans_df[loans_df['Products'] == 'PSL']
    A = employer
    B = employer_branch_book_size(employer, loans_df_filter)
    C, D = employer_branch_par(employer, 30, loans_df_filter)
    E = C - employer_branch_par(employer, 30, start_of_month_df)[0]
    F = C - employer_branch_par(employer, 30, start_of_year_df)[0]
    G = B - employer_branch_book_size(employer, start_of_month_df)
    H = B - employer_branch_book_size(employer, start_of_year_df)
    return (A, B, C, D, E, F, G, H)

def sme_generate_entry(loans_df, start_of_year_df, start_of_month_df, yesterdays_all_df, branch):
    loans_df_filter = loans_df[loans_df['Products'].isin(SME_PRODUCTS)]
    A = branch
    B = sme_branch_summ_book_size(branch, loans_df_filter)
    C, E = sme_branch_summ_par(branch, 7, loans_df_filter)
    D, F = sme_branch_summ_par(branch, 30, loans_df_filter)
    G = C - sme_branch_summ_par(branch, 7, yesterdays_all_df)[0]
    H = C - sme_branch_summ_par(branch, 7, start_of_month_df)[0]
    I = D - sme_branch_summ_par(branch, 30, yesterdays_all_df)[0]
    J = D - sme_branch_summ_par(branch, 30, start_of_month_df)[0]
    K = sme_accs_in_arrears(branch, loans_df_filter)
    if datetime.now().day == 1:
        L = 0
        M = 0
        N = 0
        O = 0
    else:
        L = loans_df[(loans_df['Days In Arrears'] < 3) & loans_df['Products'].isin(SME_PRODUCTS)]['Principal Balance'].sum()
        M = loans_df[(loans_df['Days In Arrears'] < 7) & loans_df['Products'].isin(SME_PRODUCTS)]['Principal Balance'].sum()
        N = loans_df[(loans_df['Days In Arrears'] < 14) & loans_df['Products'].isin(SME_PRODUCTS)]['Principal Balance'].sum()
        O = loans_df[(loans_df['Days In Arrears'] < 21) & loans_df['Products'].isin(SME_PRODUCTS)]['Principal Balance'].sum()
    return (A, B, C, D, E, F, G, H, I, J, K, L, M, N, O)

def sme_agent_generate_entry(loans_df, start_of_year_df, start_of_month_df, yesterdays_all_df, agent):
    loans_df_filter = loans_df[loans_df['Products'].isin(SME_PRODUCTS)]
    A = agent
    B = sme_agent_summ_book_size(agent, loans_df_filter)
    C, E = sme_agent_summ_par(agent, 7, loans_df_filter)
    D, F = sme_agent_summ_par(agent, 30, loans_df_filter)
    G = C - sme_agent_summ_par(agent, 7, yesterdays_all_df)[0]
    H = C - sme_agent_summ_par(agent, 7, start_of_month_df)[0]
    I = D - sme_agent_summ_par(agent, 30, yesterdays_all_df)[0]
    J = D - sme_agent_summ_par(agent, 30, start_of_month_df)[0]
    K = sme_agent_accs_in_arrears(agent, loans_df_filter)
    if datetime.now().day == 1:
        L = 0
        M = 0
        N = 0
        O = 0
    else:
        L = loans_df[(loans_df['Days In Arrears'] < 3) & loans_df['Products'].isin(SME_PRODUCTS)]['Principal Balance'].sum()
        M = loans_df[(loans_df['Days In Arrears'] < 7) & loans_df['Products'].isin(SME_PRODUCTS)]['Principal Balance'].sum()
        N = loans_df[(loans_df['Days In Arrears'] < 14) & loans_df['Products'].isin(SME_PRODUCTS)]['Principal Balance'].sum()
        O = loans_df[(loans_df['Days In Arrears'] < 21) & loans_df['Products'].isin(SME_PRODUCTS)]['Principal Balance'].sum()
    return (A, B, C, D, E, F, G, H, I, J, K, L, M, N, O)

def easy_phone_generate_entry(loans_df, start_of_year_df, start_of_month_df, yesterdays_all_df, sales_agent):
    loans_df_filter = loans_df[loans_df['Products'] == 'Eezy-Phone']
    start_of_year_df_filter = start_of_year_df[start_of_year_df['Products'] == 'Eezy-Phone']
    start_of_month_df_filter = start_of_month_df[start_of_month_df['Products'] == 'Eezy-Phone']
    yesterdays_all_df_filter = yesterdays_all_df[yesterdays_all_df['Products'] == 'Eezy-Phone']
    A = sales_agent
    B = loans_df_filter[loans_df_filter['Sales Representative (Loan)'] == sales_agent]['Principal Balance'].sum()
    C = easy_phone_agent_var(sales_agent, 3, loans_df_filter)
    D = easy_phone_agent_var(sales_agent, 7, loans_df_filter)
    E = easy_phone_agent_var(sales_agent, 30, loans_df_filter)
    F = agent_par(sales_agent, 3, loans_df_filter)
    G = agent_par(sales_agent, 7, loans_df_filter)
    H = agent_par(sales_agent, 30, loans_df_filter)
    I = D - easy_phone_agent_var(sales_agent, 7, yesterdays_all_df_filter)
    J = D - easy_phone_agent_var(sales_agent, 7, start_of_month_df_filter)
    K = E - easy_phone_agent_var(sales_agent, 30, yesterdays_all_df_filter)
    L = E - easy_phone_agent_var(sales_agent, 30, start_of_month_df_filter)
    M = B - yesterdays_all_df_filter[yesterdays_all_df_filter['Sales Representative (Loan)'] == sales_agent]['Principal Balance'].sum()
    N = B - start_of_month_df_filter[start_of_month_df_filter['Sales Representative (Loan)'] == sales_agent]['Principal Balance'].sum()
    O = B - start_of_year_df_filter[start_of_year_df_filter['Sales Representative (Loan)'] == sales_agent]['Principal Balance'].sum()
    return (A, B, C, D, E, F, G, H, I, J, K, L, M, N, O)

def working_capital_generate_entry(loans_df, start_of_year_df, start_of_month_df, yesterdays_all_df, sales_agent):
    loans_df_filter = loans_df[loans_df['Products'] == 'Working Capital']
    start_of_year_df_filter = start_of_year_df[start_of_year_df['Products'] == 'Working Capital']
    start_of_month_df_filter = start_of_month_df[start_of_month_df['Products'] == 'Working Capital']
    yesterdays_all_df_filter = yesterdays_all_df[yesterdays_all_df['Products'] == 'Working Capital']
    A = sales_agent
    B = loans_df_filter[loans_df_filter['Sales Representative (Loan)'] == sales_agent]['Principal Balance'].sum()
    C = easy_phone_agent_var(sales_agent, 3, loans_df_filter)
    D = easy_phone_agent_var(sales_agent, 7, loans_df_filter)
    E = easy_phone_agent_var(sales_agent, 30, loans_df_filter)
    F = agent_par(sales_agent, 3, loans_df_filter)
    G = agent_par(sales_agent, 7, loans_df_filter)
    H = agent_par(sales_agent, 30, loans_df_filter)
    I = D - easy_phone_agent_var(sales_agent, 7, yesterdays_all_df_filter)
    J = D - easy_phone_agent_var(sales_agent, 7, start_of_month_df_filter)
    K = E - easy_phone_agent_var(sales_agent, 30, yesterdays_all_df_filter)
    L = E - easy_phone_agent_var(sales_agent, 30, start_of_month_df_filter)
    M = B - yesterdays_all_df_filter[yesterdays_all_df_filter['Sales Representative (Loan)'] == sales_agent]['Principal Balance'].sum()
    N = B - start_of_month_df_filter[start_of_month_df_filter['Sales Representative (Loan)'] == sales_agent]['Principal Balance'].sum()
    O = B - start_of_year_df_filter[start_of_year_df_filter['Sales Representative (Loan)'] == sales_agent]['Principal Balance'].sum()
    return (A, B, C, D, E, F, G, H, I, J, K, L, M, N, O)

def work_through_products(loans_df, start_of_year_df, start_of_month_df, yesterdays_all_df, product, product_count, product_sheet):
    A, B, C, D, E, F, G, H, I = product_generate_entry(loans_df, start_of_year_df, start_of_month_df, yesterdays_all_df, product)
    write_entry_summary(product_sheet, A, B, C, D, E, F, G, H, I, product_count)

def work_through_collections(loans_df, start_of_year_df, start_of_month_df, yesterdays_all_df, collection_officer, collections_count, collections_sheet):
    A, B, C, D, E, F, G, H, I, J, K, L, M, N, O = collections_generate_entry(loans_df, start_of_year_df, start_of_month_df, yesterdays_all_df, collection_officer)
    write_entry_collection(collections_sheet, A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, collections_count)

def work_through_cs_branch_summary(loans_df, start_of_year_df, start_of_month_df, yesterdays_all_df, branch, branch_summary_count, branch_sheet):
    A, B, C, D, E, F, G, H, I, J, K, L, M = cs_branch_summary_generate_entry(loans_df, start_of_year_df, start_of_month_df, yesterdays_all_df, branch)
    write_entry_branch_summary(branch_sheet, A, B, C, D, E, F, G, H, I, J, K, L, M, branch_summary_count)

def work_through_lbf(loans_df, start_of_year_df, start_of_month_df, yesterdays_all_df, ro, lbf_count, lbf_sheet):
    A, B, C, D, E, F, G, H, I, J, K, L, M, N = lbf_generate_entry(loans_df, start_of_year_df, start_of_month_df, yesterdays_all_df, ro)
    write_entry_lbf(lbf_sheet, A, B, C, D, E, F, G, H, I, J, K, L, M, N, lbf_count)

def work_through_lbf_branch(loans_df, start_of_year_df, start_of_month_df, yesterdays_all_df, branch, lbf_branch_count, lbf_branch_sheet):
    A, B, C, D, E, F, G, H, I, J, K, L, M, N = lbf_branch_generate_entry(loans_df, start_of_year_df, start_of_month_df, yesterdays_all_df, branch)
    write_entry_lbf_branch(lbf_branch_sheet, A, B, C, D, E, F, G, H, I, J, K, L, M, N, lbf_branch_count)

def work_through_private_sector_employer(loans_df, start_of_year_df, start_of_month_df, yesterdays_all_df, employer, employer_count, private_sector_employer_sheet):
    A, B, C, D, E, F, G, H = private_sector_generate_entry(loans_df, start_of_year_df, start_of_month_df, yesterdays_all_df, employer)
    write_entry_private_sector_client(private_sector_employer_sheet, A, B, C, D, E, F, G, H, employer_count)

def work_through_sme(loans_df, start_of_year_df, start_of_month_df, yesterdays_all_df, branch, sme_count, sme_sheet):
    A, B, C, D, E, F, G, H, I, J, K, L, M, N, O = sme_generate_entry(loans_df, start_of_year_df, start_of_month_df, yesterdays_all_df, branch)
    write_entry_sme(sme_sheet, A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, sme_count)

def work_through_sme_agent(loans_df, start_of_year_df, start_of_month_df, yesterdays_all_df, agent, sme_agent_count, sme_sheet):
    A, B, C, D, E, F, G, H, I, J, K, L, M, N, O = sme_agent_generate_entry(loans_df, start_of_year_df, start_of_month_df, yesterdays_all_df, agent)
    write_entry_sme(sme_sheet, A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, sme_agent_count)

def work_through_easy_phone(loans_df, start_of_year_df, start_of_month_df, yesterdays_all_df, sales_agent, easy_phone_count, easy_phone_sheet):
    A, B, C, D, E, F, G, H, I, J, K, L, M, N, O = easy_phone_generate_entry(loans_df, start_of_year_df, start_of_month_df, yesterdays_all_df, sales_agent)
    write_entry_easy_phone(easy_phone_sheet, A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, easy_phone_count)

def work_through_working_capital(loans_df, start_of_year_df, start_of_month_df, yesterdays_all_df, sales_agent, working_capital_count, working_capital_sheet):
    A, B, C, D, E, F, G, H, I, J, K, L, M, N, O = working_capital_generate_entry(loans_df, start_of_year_df, start_of_month_df, yesterdays_all_df, sales_agent)
    write_entry_working_capital(working_capital_sheet, A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, working_capital_count)

def maziwa_headings(maziwa_sheet):
    maziwa_sheet['A1'] = 'Branch'
    maziwa_sheet['A1'].font = Font(size=10, bold=True)
    maziwa_sheet['A1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    maziwa_sheet['B1'] = 'Book Size'
    maziwa_sheet['B1'].font = Font(size=10, bold=True)
    maziwa_sheet['B1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    maziwa_sheet['C1'] = 'VAR7'
    maziwa_sheet['C1'].font = Font(size=10, bold=True)
    maziwa_sheet['C1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    maziwa_sheet['D1'] = 'VAR30'
    maziwa_sheet['D1'].font = Font(size=10, bold=True)
    maziwa_sheet['D1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    maziwa_sheet['E1'] = 'PAR7'
    maziwa_sheet['E1'].font = Font(size=10, bold=True)
    maziwa_sheet['E1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    maziwa_sheet['F1'] = 'PAR30'
    maziwa_sheet['F1'].font = Font(size=10, bold=True)
    maziwa_sheet['F1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    maziwa_sheet['G1'] = 'VAR7 ' + delta_icon + ' FRM Yest'
    maziwa_sheet['G1'].font = Font(size=10, bold=True)
    maziwa_sheet['G1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    maziwa_sheet['H1'] = 'VAR7 ' + delta_icon + ' FRM Beg. of Month'
    maziwa_sheet['H1'].font = Font(size=10, bold=True)
    maziwa_sheet['H1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    maziwa_sheet['I1'] = 'VAR30 ' + delta_icon + ' FRM Yest'
    maziwa_sheet['I1'].font = Font(size=10, bold=True)
    maziwa_sheet['I1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    maziwa_sheet['J1'] = 'VAR30 ' + delta_icon + ' FRM Beg. of Month'
    maziwa_sheet['J1'].font = Font(size=10, bold=True)
    maziwa_sheet['J1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    maziwa_sheet['K1'] = 'No. of accs. in Arrears'
    maziwa_sheet['K1'].font = Font(size=10, bold=True)
    maziwa_sheet['K1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    maziwa_sheet['L1'] = 'Upcoming 3 days'
    maziwa_sheet['L1'].font = Font(size=10, bold=True)
    maziwa_sheet['L1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    maziwa_sheet['M1'] = 'Upcoming 7 days'
    maziwa_sheet['M1'].font = Font(size=10, bold=True)
    maziwa_sheet['M1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    maziwa_sheet['N1'] = 'Upcoming 14 days'
    maziwa_sheet['N1'].font = Font(size=10, bold=True)
    maziwa_sheet['N1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    maziwa_sheet['O1'] = 'Upcoming 21 days'
    maziwa_sheet['O1'].font = Font(size=10, bold=True)
    maziwa_sheet['O1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')

def maziwa_agents_headings(maziwa_agent_sheet):
    maziwa_agent_sheet['A1'] = 'Loan Officers'
    maziwa_agent_sheet['A1'].font = Font(size=10, bold=True)
    maziwa_agent_sheet['A1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    maziwa_agent_sheet['B1'] = 'Book Size'
    maziwa_agent_sheet['B1'].font = Font(size=10, bold=True)
    maziwa_agent_sheet['B1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    maziwa_agent_sheet['C1'] = 'VAR7'
    maziwa_agent_sheet['C1'].font = Font(size=10, bold=True)
    maziwa_agent_sheet['C1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    maziwa_agent_sheet['D1'] = 'VAR30'
    maziwa_agent_sheet['D1'].font = Font(size=10, bold=True)
    maziwa_agent_sheet['D1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    maziwa_agent_sheet['E1'] = 'PAR7'
    maziwa_agent_sheet['E1'].font = Font(size=10, bold=True)
    maziwa_agent_sheet['E1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    maziwa_agent_sheet['F1'] = 'PAR30'
    maziwa_agent_sheet['F1'].font = Font(size=10, bold=True)
    maziwa_agent_sheet['F1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    maziwa_agent_sheet['G1'] = 'VAR7 ' + delta_icon + ' FRM Yest'
    maziwa_agent_sheet['G1'].font = Font(size=10, bold=True)
    maziwa_agent_sheet['G1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    maziwa_agent_sheet['H1'] = 'VAR7 ' + delta_icon + ' FRM Beg. of Month'
    maziwa_agent_sheet['H1'].font = Font(size=10, bold=True)
    maziwa_agent_sheet['H1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    maziwa_agent_sheet['I1'] = 'VAR30 ' + delta_icon + ' FRM Yest'
    maziwa_agent_sheet['I1'].font = Font(size=10, bold=True)
    maziwa_agent_sheet['I1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    maziwa_agent_sheet['J1'] = 'VAR30 ' + delta_icon + ' FRM Beg. of Month'
    maziwa_agent_sheet['J1'].font = Font(size=10, bold=True)
    maziwa_agent_sheet['J1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    maziwa_agent_sheet['K1'] = 'No. of accs. in Arrears'
    maziwa_agent_sheet['K1'].font = Font(size=10, bold=True)
    maziwa_agent_sheet['K1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    maziwa_agent_sheet['L1'] = 'Upcoming 3 days'
    maziwa_agent_sheet['L1'].font = Font(size=10, bold=True)
    maziwa_agent_sheet['L1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    maziwa_agent_sheet['M1'] = 'Upcoming 7 days'
    maziwa_agent_sheet['M1'].font = Font(size=10, bold=True)
    maziwa_agent_sheet['M1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    maziwa_agent_sheet['N1'] = 'Upcoming 14 days'
    maziwa_agent_sheet['N1'].font = Font(size=10, bold=True)
    maziwa_agent_sheet['N1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')
    maziwa_agent_sheet['O1'] = 'Upcoming 21 days'
    maziwa_agent_sheet['O1'].font = Font(size=10, bold=True)
    maziwa_agent_sheet['O1'].fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.Color(rgb='0FEEFF'), fill_type='solid')

def write_entry_maziwa(maziwa_sheet, A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, row):
    maziwa_sheet['A' + str(row)] = A
    maziwa_sheet['A' + str(row)].font = Font(size=9)
    maziwa_sheet['A' + str(row)].number_format = '@'
    maziwa_sheet['B' + str(row)] = B
    maziwa_sheet['B' + str(row)].font = Font(size=9)
    maziwa_sheet['B' + str(row)].number_format = '#,##0'
    maziwa_sheet['C' + str(row)] = C
    maziwa_sheet['C' + str(row)].font = Font(size=9)
    maziwa_sheet['C' + str(row)].number_format = '#,##0'
    maziwa_sheet['D' + str(row)] = D
    maziwa_sheet['D' + str(row)].font = Font(size=9)
    maziwa_sheet['D' + str(row)].number_format = '#,##0'
    maziwa_sheet['E' + str(row)] = E
    maziwa_sheet['E' + str(row)].font = Font(size=9)
    maziwa_sheet['E' + str(row)].number_format = '0.00%'
    maziwa_sheet['F' + str(row)] = F
    maziwa_sheet['F' + str(row)].font = Font(size=9)
    maziwa_sheet['F' + str(row)].number_format = '0.00%'
    maziwa_sheet['G' + str(row)] = G
    maziwa_sheet['G' + str(row)].font = Font(size=9)
    maziwa_sheet['G' + str(row)].number_format = '#,##0'
    maziwa_sheet['H' + str(row)] = H
    maziwa_sheet['H' + str(row)].font = Font(size=9)
    maziwa_sheet['H' + str(row)].number_format = '#,##0'
    maziwa_sheet['I' + str(row)] = I
    maziwa_sheet['I' + str(row)].font = Font(size=9)
    maziwa_sheet['I' + str(row)].number_format = '#,##0'
    maziwa_sheet['J' + str(row)] = J
    maziwa_sheet['J' + str(row)].font = Font(size=9)
    maziwa_sheet['J' + str(row)].number_format = '#,##0'
    maziwa_sheet['K' + str(row)] = K
    maziwa_sheet['K' + str(row)].font = Font(size=9)
    maziwa_sheet['K' + str(row)].number_format = '#,##0'
    maziwa_sheet['L' + str(row)] = L
    maziwa_sheet['L' + str(row)].font = Font(size=9)
    maziwa_sheet['L' + str(row)].number_format = '#,##0'
    maziwa_sheet['M' + str(row)] = M
    maziwa_sheet['M' + str(row)].font = Font(size=9)
    maziwa_sheet['M' + str(row)].number_format = '#,##0'
    maziwa_sheet['N' + str(row)] = N
    maziwa_sheet['N' + str(row)].font = Font(size=9)
    maziwa_sheet['N' + str(row)].number_format = '#,##0'
    maziwa_sheet['O' + str(row)] = O
    maziwa_sheet['O' + str(row)].font = Font(size=9)
    maziwa_sheet['O' + str(row)].number_format = '#,##0'

def write_entry_maziwa_agent(maziwa_agent_sheet, A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, row):
    maziwa_agent_sheet['A' + str(row)] = A
    maziwa_agent_sheet['A' + str(row)].font = Font(size=9)
    maziwa_agent_sheet['A' + str(row)].number_format = '@'
    maziwa_agent_sheet['B' + str(row)] = B
    maziwa_agent_sheet['B' + str(row)].font = Font(size=9)
    maziwa_agent_sheet['B' + str(row)].number_format = '#,##0'
    maziwa_agent_sheet['C' + str(row)] = C
    maziwa_agent_sheet['C' + str(row)].font = Font(size=9)
    maziwa_agent_sheet['C' + str(row)].number_format = '#,##0'
    maziwa_agent_sheet['D' + str(row)] = D
    maziwa_agent_sheet['D' + str(row)].font = Font(size=9)
    maziwa_agent_sheet['D' + str(row)].number_format = '#,##0'
    maziwa_agent_sheet['E' + str(row)] = E
    maziwa_agent_sheet['E' + str(row)].font = Font(size=9)
    maziwa_agent_sheet['E' + str(row)].number_format = '0.00%'
    maziwa_agent_sheet['F' + str(row)] = F
    maziwa_agent_sheet['F' + str(row)].font = Font(size=9)
    maziwa_agent_sheet['F' + str(row)].number_format = '0.00%'
    maziwa_agent_sheet['G' + str(row)] = G
    maziwa_agent_sheet['G' + str(row)].font = Font(size=9)
    maziwa_agent_sheet['G' + str(row)].number_format = '#,##0'
    maziwa_agent_sheet['H' + str(row)] = H
    maziwa_agent_sheet['H' + str(row)].font = Font(size=9)
    maziwa_agent_sheet['H' + str(row)].number_format = '#,##0'
    maziwa_agent_sheet['I' + str(row)] = I
    maziwa_agent_sheet['I' + str(row)].font = Font(size=9)
    maziwa_agent_sheet['I' + str(row)].number_format = '#,##0'
    maziwa_agent_sheet['J' + str(row)] = J
    maziwa_agent_sheet['J' + str(row)].font = Font(size=9)
    maziwa_agent_sheet['J' + str(row)].number_format = '#,##0'
    maziwa_agent_sheet['K' + str(row)] = K
    maziwa_agent_sheet['K' + str(row)].font = Font(size=9)
    maziwa_agent_sheet['K' + str(row)].number_format = '#,##0'
    maziwa_agent_sheet['L' + str(row)] = L
    maziwa_agent_sheet['L' + str(row)].font = Font(size=9)
    maziwa_agent_sheet['L' + str(row)].number_format = '#,##0'
    maziwa_agent_sheet['M' + str(row)] = M
    maziwa_agent_sheet['M' + str(row)].font = Font(size=9)
    maziwa_agent_sheet['M' + str(row)].number_format = '#,##0'
    maziwa_agent_sheet['N' + str(row)] = N
    maziwa_agent_sheet['N' + str(row)].font = Font(size=9)
    maziwa_agent_sheet['N' + str(row)].number_format = '#,##0'
    maziwa_agent_sheet['O' + str(row)] = O
    maziwa_agent_sheet['O' + str(row)].font = Font(size=9)
    maziwa_agent_sheet['O' + str(row)].number_format = '#,##0'

def maziwa_branch_summ_book_size(branch, loans_df):
    branch_filter_df = loans_df[(loans_df['Branch'] == branch) & (loans_df['Products'] == 'MAZIWA')]
    total_portfolio = branch_filter_df['Principal Balance'].sum()
    return total_portfolio

def maziwa_branch_summ_par(branch, days, loans_df):
    branch_filter_df = loans_df[(loans_df['Branch'] == branch) & (loans_df['Products'] == 'MAZIWA')]
    value_in_arr = branch_filter_df[branch_filter_df['Days In Arrears'] > days]['Principal Balance'].sum()
    portfolio = branch_filter_df['Principal Balance'].sum()
    par = value_in_arr / portfolio if portfolio > 0 else 0
    return (value_in_arr, par)

def maziwa_branch_par_mv(branch, loans_df):
    branch_filter_df = loans_df[loans_df['Branch'] == branch]
    value_in_arr = branch_filter_df[branch_filter_df['Days In Arrears'] > 30]['Principal Balance'].sum()
    principal_bal = branch_filter_df['Principal Balance'].sum()
    mvmt = value_in_arr - 0.03 * principal_bal
    if mvmt > 0:
        return mvmt
    return None

def maziwa_accs_in_arrears(branch, loans_df):
    branch_filter_df = loans_df[(loans_df['Branch'] == branch) & (loans_df['Products'] == 'MAZIWA')]
    accs_in_arrs = branch_filter_df[branch_filter_df['Account State'] == 'In Arrears'].shape[0]
    return accs_in_arrs

def maziwa_agent_summ_book_size(agent, loans_df):
    branch_filter_df = loans_df[(loans_df['Sales Representative (Loan)'] == agent) & (loans_df['Products'] == 'MAZIWA')]
    total_portfolio = branch_filter_df['Principal Balance'].sum()
    return total_portfolio

def maziwa_agent_summ_par(agent, days, loans_df):
    branch_filter_df = loans_df[(loans_df['Sales Representative (Loan)'] == agent) & (loans_df['Products'] == 'MAZIWA')]
    value_in_arr = branch_filter_df[branch_filter_df['Days In Arrears'] > days]['Principal Balance'].sum()
    portfolio = branch_filter_df['Principal Balance'].sum()
    par = value_in_arr / portfolio if portfolio > 0 else 0
    return (value_in_arr, par)

def maziwa_agent_par_mv(agent, loans_df):
    branch_filter_df = loans_df[loans_df['Sales Representative (Loan)'] == agent]
    value_in_arr = branch_filter_df[branch_filter_df['Days In Arrears'] > 30]['Principal Balance'].sum()
    principal_bal = branch_filter_df['Principal Balance'].sum()
    mvmt = value_in_arr - 0.03 * principal_bal
    if mvmt > 0:
        return mvmt
    return None

def maziwa_agent_accs_in_arrears(agent, loans_df):
    branch_filter_df = loans_df[(loans_df['Sales Representative (Loan)'] == agent) & (loans_df['Products'] == 'MAZIWA')]
    accs_in_arrs = branch_filter_df[branch_filter_df['Account State'] == 'In Arrears'].shape[0]
    return accs_in_arrs

def maziwa_generate_entry(loans_df, start_of_year_df, start_of_month_df, yesterdays_all_df, branch):
    loans_df_filter = loans_df[loans_df['Products'] == 'MAZIWA']
    A = branch
    B = maziwa_branch_summ_book_size(branch, loans_df_filter)
    C, E = maziwa_branch_summ_par(branch, 7, loans_df_filter)
    D, F = maziwa_branch_summ_par(branch, 30, loans_df_filter)
    G = C - maziwa_branch_summ_par(branch, 7, yesterdays_all_df)[0]
    H = C - maziwa_branch_summ_par(branch, 7, start_of_month_df)[0]
    I = D - maziwa_branch_summ_par(branch, 30, yesterdays_all_df)[0]
    J = D - maziwa_branch_summ_par(branch, 30, start_of_month_df)[0]
    K = maziwa_accs_in_arrears(branch, loans_df_filter)
    if datetime.now().day == 1:
        L = 0
        M = 0
        N = 0
        O = 0
    else:
        L = loans_df[(loans_df['Days In Arrears'] < 3) & (loans_df['Products'] == 'MAZIWA')]['Principal Balance'].sum()
        M = loans_df[(loans_df['Days In Arrears'] < 7) & (loans_df['Products'] == 'MAZIWA')]['Principal Balance'].sum()
        N = loans_df[(loans_df['Days In Arrears'] < 14) & (loans_df['Products'] == 'MAZIWA')]['Principal Balance'].sum()
        O = loans_df[(loans_df['Days In Arrears'] < 21) & (loans_df['Products'] == 'MAZIWA')]['Principal Balance'].sum()
    return (A, B, C, D, E, F, G, H, I, J, K, L, M, N, O)

def maziwa_agent_generate_entry(loans_df, start_of_year_df, start_of_month_df, yesterdays_all_df, agent):
    loans_df_filter = loans_df[loans_df['Products'] == 'MAZIWA']
    A = agent
    B = maziwa_agent_summ_book_size(agent, loans_df_filter)
    C, E = maziwa_agent_summ_par(agent, 7, loans_df_filter)
    D, F = maziwa_agent_summ_par(agent, 30, loans_df_filter)
    G = C - maziwa_agent_summ_par(agent, 7, yesterdays_all_df)[0]
    H = C - maziwa_agent_summ_par(agent, 7, start_of_month_df)[0]
    I = D - maziwa_agent_summ_par(agent, 30, yesterdays_all_df)[0]
    J = D - maziwa_agent_summ_par(agent, 30, start_of_month_df)[0]
    K = maziwa_agent_accs_in_arrears(agent, loans_df_filter)
    if datetime.now().day == 1:
        L = 0
        M = 0
        N = 0
        O = 0
    else:
        L = loans_df[(loans_df['Days In Arrears'] < 3) & (loans_df['Products'] == 'MAZIWA')]['Principal Balance'].sum()
        M = loans_df[(loans_df['Days In Arrears'] < 7) & (loans_df['Products'] == 'MAZIWA')]['Principal Balance'].sum()
        N = loans_df[(loans_df['Days In Arrears'] < 14) & (loans_df['Products'] == 'MAZIWA')]['Principal Balance'].sum()
        O = loans_df[(loans_df['Days In Arrears'] < 21) & (loans_df['Products'] == 'MAZIWA')]['Principal Balance'].sum()
    return (A, B, C, D, E, F, G, H, I, J, K, L, M, N, O)

def work_through_maziwa(loans_df, start_of_year_df, start_of_month_df, yesterdays_all_df, branch, maziwa_count, maziwa_sheet):
    A, B, C, D, E, F, G, H, I, J, K, L, M, N, O = maziwa_generate_entry(loans_df, start_of_year_df, start_of_month_df, yesterdays_all_df, branch)
    write_entry_maziwa(maziwa_sheet, A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, maziwa_count)

def work_through_maziwa_agent(loans_df, start_of_year_df, start_of_month_df, yesterdays_all_df, agent, maziwa_agent_count, maziwa_sheet):
    A, B, C, D, E, F, G, H, I, J, K, L, M, N, O = maziwa_agent_generate_entry(loans_df, start_of_year_df, start_of_month_df, yesterdays_all_df, agent)
    write_entry_maziwa(maziwa_sheet, A, B, C, D, E, F, G, H, I, J, K, L, M, N, O, maziwa_agent_count)

def main():
    loans_file = input("Please input today's loan listing... \n")
    loans_df = pd.read_excel(loans_file)
    yesterdays_loans_file = input("Please input the beginning of the yesterday's loan listing... \n")
    yesterdays_loans_df = pd.read_excel(yesterdays_loans_file)
    beg_of_month_loans_file = input("Please input the beginning of the month's loan listing... \n")
    beg_of_month_loans_df = pd.read_excel(beg_of_month_loans_file)
    beg_of_year_loans_file = input("Please input the beginning of the year's loan listing... \n")
    beg_of_year_loans_df = pd.read_excel(beg_of_year_loans_file)
    output_book = openpyxl.Workbook()
    summary_sheet = output_book.active
    summary_sheet.title = 'Summary'
    summary_headings(summary_sheet)
    collections_sheet = output_book.create_sheet('CS collection Officer')
    collections_headings(collections_sheet)
    cs_branch_summary_sheet = output_book.create_sheet('CS Branch BPs Summary')
    cs_branches_summary_headings(cs_branch_summary_sheet)
    ro_sheet = output_book.create_sheet('LBF ROs')
    lbf_headings(ro_sheet)
    lbf_branch_sheet = output_book.create_sheet('LBF Branches')
    lbf_branches_headings(lbf_branch_sheet)
    private_sector_sheet = output_book.create_sheet('Private Sector Employer')
    private_sector_employer_headings(private_sector_sheet)
    sme_sheet = output_book.create_sheet('SME Branches')
    sme_headings(sme_sheet)
    sme_agent_sheet = output_book.create_sheet('SME Loan Officers')
    sme_agents_headings(sme_agent_sheet)
    maziwa_sheet = output_book.create_sheet('Maziwa Branches')
    maziwa_headings(maziwa_sheet)
    maziwa_agent_sheet = output_book.create_sheet('Maziwa Loan Officers')
    maziwa_agents_headings(maziwa_agent_sheet)
    easy_phone_sheet = output_book.create_sheet('Easy Phone')
    easy_phone_headings(easy_phone_sheet)
    working_capital_sheet = output_book.create_sheet('Working Capital Wallet')
    easy_phone_headings(working_capital_sheet)
    first = FormatObject(type='num', val=-1e-18)
    mid = FormatObject(type='num', val=0)
    last = FormatObject(type='num', val=1e-18)
    colors = [Color('e6b8b7'), Color('fcd6b4'), Color('d8e4bc')]
    colors_inv = [Color('d8e4bc'), Color('fcd6b4'), Color('e6b8b7')]
    color_scale = ColorScale(cfvo=[first, mid, last], color=colors)
    color_scale_inv = ColorScale(cfvo=[first, mid, last], color=colors_inv)
    rule = Rule(type='colorScale', colorScale=color_scale)
    rule_inv = Rule(type='colorScale', colorScale=color_scale_inv)
    collections_sheet.conditional_formatting.add('K2:K20000', rule_inv)
    collections_sheet.conditional_formatting.add('L2:L20000', rule_inv)
    collections_sheet.conditional_formatting.add('M2:M20000', rule_inv)
    collections_sheet.conditional_formatting.add('N2:N20000', rule)
    collections_sheet.conditional_formatting.add('O2:O20000', rule)
    collections_sheet.freeze_panes = collections_sheet['B2']
    cs_branch_summary_sheet.conditional_formatting.add('L2:L20000', rule_inv)
    cs_branch_summary_sheet.conditional_formatting.add('M2:M20000', rule_inv)
    cs_branch_summary_sheet.freeze_panes = cs_branch_summary_sheet['B2']
    ro_sheet.conditional_formatting.add('J2:J20000', rule_inv)
    ro_sheet.conditional_formatting.add('K2:K20000', rule_inv)
    ro_sheet.conditional_formatting.add('L2:L20000', rule_inv)
    ro_sheet.conditional_formatting.add('M2:M20000', rule)
    ro_sheet.conditional_formatting.add('N2:N20000', rule)
    ro_sheet.freeze_panes = ro_sheet['B2']
    lbf_branch_sheet.conditional_formatting.add('J2:J20000', rule_inv)
    lbf_branch_sheet.conditional_formatting.add('K2:K20000', rule_inv)
    lbf_branch_sheet.conditional_formatting.add('L2:L20000', rule_inv)
    lbf_branch_sheet.conditional_formatting.add('M2:M20000', rule)
    lbf_branch_sheet.conditional_formatting.add('N2:N20000', rule)
    lbf_branch_sheet.freeze_panes = lbf_branch_sheet['B2']
    private_sector_sheet.conditional_formatting.add('E2:E20000', rule_inv)
    private_sector_sheet.conditional_formatting.add('F2:F20000', rule_inv)
    private_sector_sheet.conditional_formatting.add('G2:G20000', rule)
    private_sector_sheet.conditional_formatting.add('H2:H20000', rule)
    private_sector_sheet.freeze_panes = private_sector_sheet['B2']
    sme_sheet.conditional_formatting.add('G2:G20000', rule_inv)
    sme_sheet.conditional_formatting.add('H2:H20000', rule_inv)
    sme_sheet.conditional_formatting.add('I2:I20000', rule_inv)
    sme_sheet.conditional_formatting.add('J2:J20000', rule_inv)
    sme_sheet.freeze_panes = sme_sheet['B2']
    sme_agent_sheet.conditional_formatting.add('G2:G20000', rule_inv)
    sme_agent_sheet.conditional_formatting.add('H2:H20000', rule_inv)
    sme_agent_sheet.conditional_formatting.add('I2:I20000', rule_inv)
    sme_agent_sheet.conditional_formatting.add('J2:J20000', rule_inv)
    sme_agent_sheet.freeze_panes = sme_agent_sheet['B2']
    maziwa_sheet.conditional_formatting.add('G2:G20000', rule_inv)
    maziwa_sheet.conditional_formatting.add('H2:H20000', rule_inv)
    maziwa_sheet.conditional_formatting.add('I2:I20000', rule_inv)
    maziwa_sheet.conditional_formatting.add('J2:J20000', rule_inv)
    maziwa_sheet.freeze_panes = maziwa_sheet['B2']
    maziwa_agent_sheet.conditional_formatting.add('G2:G20000', rule_inv)
    maziwa_agent_sheet.conditional_formatting.add('H2:H20000', rule_inv)
    maziwa_agent_sheet.conditional_formatting.add('I2:I20000', rule_inv)
    maziwa_agent_sheet.conditional_formatting.add('J2:J20000', rule_inv)
    maziwa_agent_sheet.freeze_panes = maziwa_agent_sheet['B2']
    easy_phone_sheet.conditional_formatting.add('I2:I20000', rule_inv)
    easy_phone_sheet.conditional_formatting.add('J2:J20000', rule_inv)
    easy_phone_sheet.conditional_formatting.add('K2:K20000', rule_inv)
    easy_phone_sheet.conditional_formatting.add('L2:L20000', rule_inv)
    easy_phone_sheet.conditional_formatting.add('M2:M20000', rule)
    easy_phone_sheet.conditional_formatting.add('N2:N20000', rule)
    easy_phone_sheet.conditional_formatting.add('O2:O20000', rule)
    easy_phone_sheet.freeze_panes = easy_phone_sheet['B2']
    working_capital_sheet.conditional_formatting.add('I2:I20000', rule_inv)
    working_capital_sheet.conditional_formatting.add('J2:J20000', rule_inv)
    working_capital_sheet.conditional_formatting.add('K2:K20000', rule_inv)
    working_capital_sheet.conditional_formatting.add('L2:L20000', rule_inv)
    working_capital_sheet.conditional_formatting.add('M2:M20000', rule)
    working_capital_sheet.conditional_formatting.add('N2:N20000', rule)
    working_capital_sheet.conditional_formatting.add('O2:O20000', rule)
    working_capital_sheet.freeze_panes = working_capital_sheet['B2']
    products = loans_df[loans_df['Products'].notna()].groupby('Products')['Principal Balance'].sum().reset_index().sort_values(by='Principal Balance', ascending=False)['Products'].unique()
    collection_officers = loans_df[(loans_df['Products'] == 'CSL') & loans_df['Collection_Officer'].notna()].groupby('Collection_Officer')['Principal Balance'].sum().reset_index().sort_values(by='Principal Balance', ascending=False)['Collection_Officer'].unique()
    cs_branch_list = loans_df[loans_df['Products'] == 'CSL'].groupby('Branch')['Principal Balance'].sum().reset_index().sort_values(by='Principal Balance', ascending=False)['Branch'].unique().tolist()
    relationship_managers = loans_df[loans_df['Relationship/Recoveries Officers'].notna()].groupby('Relationship/Recoveries Officers')['Principal Balance'].sum().reset_index().sort_values(by='Principal Balance', ascending=False)['Relationship/Recoveries Officers'].unique()
    lbf_branches = loans_df[loans_df['Products'] == 'LBF'].groupby('Branch')['Principal Balance'].sum().reset_index().sort_values(by='Principal Balance', ascending=False)['Branch'].unique().tolist()
    private_sector_employers = loans_df[(loans_df['Products'] == 'PSL') & loans_df['Employer_Name (Client)'].notna()].groupby('Employer_Name (Client)')['Principal Balance'].sum().reset_index().sort_values(by='Principal Balance', ascending=False)['Employer_Name (Client)'].unique().tolist()
    sme_branches = loans_df[loans_df['Products'].isin(SME_PRODUCTS)].groupby('Branch')['Principal Balance'].sum().reset_index().sort_values(by='Principal Balance', ascending=False)['Branch'].unique().tolist()
    sme_loan_officers = loans_df[loans_df['Products'].isin(SME_PRODUCTS) & loans_df['Sales Representative (Loan)'].notna()].groupby('Sales Representative (Loan)')['Principal Balance'].sum().reset_index().sort_values(by='Principal Balance', ascending=False)['Sales Representative (Loan)'].unique().tolist()
    maziwa_branches = loans_df[loans_df['Products'] == 'MAZIWA'].groupby('Branch')['Principal Balance'].sum().reset_index().sort_values(by='Principal Balance', ascending=False)['Branch'].unique().tolist()
    maziwa_loan_officers = loans_df[(loans_df['Products'] == 'MAZIWA') & loans_df['Sales Representative (Loan)'].notna()].groupby('Sales Representative (Loan)')['Principal Balance'].sum().reset_index().sort_values(by='Principal Balance', ascending=False)['Sales Representative (Loan)'].unique().tolist()
    easy_phone_agents = loans_df[(loans_df['Products'] == 'Eezy-Phone') & loans_df['Sales Representative (Loan)'].notna()].groupby('Sales Representative (Loan)')['Principal Balance'].sum().reset_index().sort_values(by='Principal Balance', ascending=False)['Sales Representative (Loan)'].unique().tolist()
    working_capital_agents = loans_df[(loans_df['Products'] == 'Working Capital') & loans_df['Sales Representative (Loan)'].notna()].groupby('Sales Representative (Loan)')['Principal Balance'].sum().reset_index().sort_values(by='Principal Balance', ascending=False)['Sales Representative (Loan)'].unique().tolist()
    summary_sheet.sheet_view.showGridLines = False
    collections_sheet.sheet_view.showGridLines = False
    cs_branch_summary_sheet.sheet_view.showGridLines = False
    ro_sheet.sheet_view.showGridLines = False
    lbf_branch_sheet.sheet_view.showGridLines = False
    private_sector_sheet.sheet_view.showGridLines = False
    sme_sheet.sheet_view.showGridLines = False
    sme_agent_sheet.sheet_view.showGridLines = False
    maziwa_sheet.sheet_view.showGridLines = False
    maziwa_agent_sheet.sheet_view.showGridLines = False
    easy_phone_sheet.sheet_view.showGridLines = False
    working_capital_sheet.sheet_view.showGridLines = False
    product_count = 2
    print('Work Through Products')
    for product in tqdm(products, colour='green'):
        work_through_products(loans_df, beg_of_year_loans_df, beg_of_month_loans_df, yesterdays_loans_df, product, product_count, summary_sheet)
        summary_sheet.conditional_formatting.add('E2:E' + str(product_count), rule_inv)
        summary_sheet.conditional_formatting.add('F2:F' + str(product_count), rule_inv)
        summary_sheet.conditional_formatting.add('G2:G' + str(product_count), rule_inv)
        summary_sheet.conditional_formatting.add('H2:H' + str(product_count), rule)
        summary_sheet.conditional_formatting.add('I2:I' + str(product_count), rule)
        product_count += 1
    set_border(summary_sheet, 'A2:I' + str(product_count - 1))
    summary_sheet['A' + str(product_count)] = 'Total Book Size'
    summary_sheet['A' + str(product_count)].font = Font(size=10, bold=True)
    summary_sheet['A' + str(product_count)].number_format = '@'
    summary_sheet['B' + str(product_count)].font = Font(size=10, bold=True)
    summary_sheet['B' + str(product_count)].value = loans_df['Principal Balance'].sum()
    summary_sheet['B' + str(product_count)].number_format = '#,##0'
    summary_sheet['C' + str(product_count)].font = Font(size=10, bold=True)
    summary_sheet['C' + str(product_count)].value = loans_df[loans_df['Days In Arrears'] > 30]['Principal Balance'].sum()
    summary_sheet['C' + str(product_count)].number_format = '#,##0'
    summary_sheet['D' + str(product_count)].font = Font(size=10, bold=True)
    summary_sheet['D' + str(product_count)].value = loans_df[loans_df['Days In Arrears'] > 30]['Principal Balance'].sum() / loans_df['Principal Balance'].sum() if loans_df['Principal Balance'].sum() else 0
    summary_sheet['D' + str(product_count)].number_format = '0.00%'
    summary_sheet['E' + str(product_count)].font = Font(size=10, bold=True)
    summary_sheet['E' + str(product_count)].value = loans_df[loans_df['Days In Arrears'] > 30]['Principal Balance'].sum() - yesterdays_loans_df[yesterdays_loans_df['Days In Arrears'] > 30]['Principal Balance'].sum()
    summary_sheet['E' + str(product_count)].number_format = '#,##0'
    summary_sheet['F' + str(product_count)].font = Font(size=10, bold=True)
    summary_sheet['F' + str(product_count)].value = loans_df[loans_df['Days In Arrears'] > 30]['Principal Balance'].sum() - beg_of_month_loans_df[beg_of_month_loans_df['Days In Arrears'] > 30]['Principal Balance'].sum()
    summary_sheet['F' + str(product_count)].number_format = '#,##0'
    summary_sheet['G' + str(product_count)].font = Font(size=10, bold=True)
    summary_sheet['G' + str(product_count)].value = loans_df[loans_df['Days In Arrears'] > 30]['Principal Balance'].sum() - beg_of_year_loans_df[beg_of_year_loans_df['Days In Arrears'] > 30]['Principal Balance'].sum()
    summary_sheet['G' + str(product_count)].number_format = '#,##0'
    summary_sheet['H' + str(product_count)].font = Font(size=10, bold=True)
    summary_sheet['H' + str(product_count)].value = loans_df['Principal Balance'].sum() - beg_of_month_loans_df['Principal Balance'].sum()
    summary_sheet['H' + str(product_count)].number_format = '#,##0'
    summary_sheet['I' + str(product_count)].font = Font(size=10, bold=True)
    summary_sheet['I' + str(product_count)].value = loans_df['Principal Balance'].sum() - beg_of_year_loans_df['Principal Balance'].sum()
    summary_sheet['I' + str(product_count)].number_format = '#,##0'
    summary_sheet.conditional_formatting.add('E2:E' + str(product_count), rule_inv)
    summary_sheet.conditional_formatting.add('F2:F' + str(product_count), rule_inv)
    summary_sheet.conditional_formatting.add('G2:G' + str(product_count), rule_inv)
    summary_sheet.conditional_formatting.add('H2:H' + str(product_count), rule)
    summary_sheet.conditional_formatting.add('I2:I' + str(product_count), rule)
    product_count += 1
    by_product_count = product_count + 2
    by_products = loans_df[loans_df['Collections Status'].notna()].groupby(['Collections Status'])['Principal Balance'].agg('count').reset_index()
    by_products = by_products.sort_values(by=['Principal Balance'], ascending=False)
    by_products = by_products['Collections Status'].unique()
    summary_sheet['A' + str(by_product_count)] = 'CS Loans BP Status Summary'
    summary_sheet['A' + str(by_product_count)].alignment = Alignment(horizontal='center')
    summary_sheet['A' + str(by_product_count)].font = Font(size=10, bold=True)
    summary_sheet['B' + str(by_product_count)] = 'Count'
    summary_sheet['B' + str(by_product_count)].alignment = Alignment(horizontal='center')
    summary_sheet['B' + str(by_product_count)].font = Font(size=10, bold=True)
    summary_sheet['C' + str(by_product_count)] = 'Book Size'
    summary_sheet['C' + str(by_product_count)].alignment = Alignment(horizontal='center')
    summary_sheet['C' + str(by_product_count)].font = Font(size=10, bold=True)
    summary_sheet['D' + str(by_product_count)] = delta_icon + ' from Beg. of Month(Count)'
    summary_sheet['D' + str(by_product_count)].alignment = Alignment(horizontal='center')
    summary_sheet['D' + str(by_product_count)].font = Font(size=10, bold=True)
    summary_sheet['E' + str(by_product_count)] = delta_icon + ' from Beg. of Month(Portfolio)'
    summary_sheet['E' + str(by_product_count)].alignment = Alignment(horizontal='center')
    summary_sheet['E' + str(by_product_count)].font = Font(size=10, bold=True)
    summary_sheet['F' + str(by_product_count)] = delta_icon + ' from Beg. of Year(Count)'
    summary_sheet['F' + str(by_product_count)].alignment = Alignment(horizontal='center')
    summary_sheet['F' + str(by_product_count)].font = Font(size=10, bold=True)
    summary_sheet['G' + str(by_product_count)] = delta_icon + ' from Beg. of Year(Portfolio)'
    summary_sheet['G' + str(by_product_count)].alignment = Alignment(horizontal='center')
    summary_sheet['G' + str(by_product_count)].font = Font(size=10, bold=True)
    by_product_border_count = by_product_count
    by_product_count += 1
    for by_product in by_products:
        A = by_product
        if by_product == 'Performing/ Normal':
            B = loans_df[(loans_df['Products'] == 'CSL') & (loans_df['Collections Status'] == by_product) | loans_df['Collections Status'].isna() & (loans_df['Account State'] == 'Active') & (loans_df['Products'] == 'CSL')].shape[0]
            C = loans_df[(loans_df['Products'] == 'CSL') & (loans_df['Collections Status'] == by_product) | loans_df['Collections Status'].isna() & (loans_df['Account State'] == 'Active') & (loans_df['Products'] == 'CSL')]['Principal Balance'].sum()
            D = loans_df[(loans_df['Products'] == 'CSL') & (loans_df['Collections Status'] == by_product) | loans_df['Collections Status'].isna() & (loans_df['Account State'] == 'Active') & (loans_df['Products'] == 'CSL')].shape[0] - beg_of_month_loans_df[(beg_of_month_loans_df['Products'] == 'CSL') & (beg_of_month_loans_df['Collections Status'] == by_product) | beg_of_month_loans_df['Collections Status'].isna() & (beg_of_month_loans_df['Account State'] == 'Active') & (beg_of_month_loans_df['Products'] == 'CSL')].shape[0]
            E = loans_df[(loans_df['Products'] == 'CSL') & (loans_df['Collections Status'] == by_product) | loans_df['Collections Status'].isna() & (loans_df['Account State'] == 'Active') & (loans_df['Products'] == 'CSL')]['Principal Balance'].sum() - beg_of_month_loans_df[(beg_of_month_loans_df['Products'] == 'CSL') & (beg_of_month_loans_df['Collections Status'] == by_product) | beg_of_month_loans_df['Collections Status'].isna() & (beg_of_month_loans_df['Account State'] == 'Active') & (beg_of_month_loans_df['Products'] == 'CSL')]['Principal Balance'].sum()
            F = loans_df[(loans_df['Products'] == 'CSL') & (loans_df['Collections Status'] == by_product) | loans_df['Collections Status'].isna() & (loans_df['Account State'] == 'Active') & (loans_df['Products'] == 'CSL')].shape[0] - beg_of_year_loans_df[(beg_of_year_loans_df['Products'] == 'CSL') & (beg_of_year_loans_df['Collections Status'] == by_product) | beg_of_year_loans_df['Collections Status'].isna() & (beg_of_year_loans_df['Account State'] == 'Active') & (beg_of_year_loans_df['Products'] == 'CSL')].shape[0]
            G = loans_df[(loans_df['Products'] == 'CSL') & (loans_df['Collections Status'] == by_product) | loans_df['Collections Status'].isna() & (loans_df['Account State'] == 'Active') & (loans_df['Products'] == 'CSL')]['Principal Balance'].sum() - beg_of_year_loans_df[(beg_of_year_loans_df['Products'] == 'CSL') & (beg_of_year_loans_df['Collections Status'] == by_product) | beg_of_year_loans_df['Collections Status'].isna() & (beg_of_year_loans_df['Account State'] == 'Active') & (beg_of_year_loans_df['Products'] == 'CSL')]['Principal Balance'].sum()
        else:
            B = loans_df[loans_df['Collections Status'] == by_product].shape[0]
            C = loans_df[loans_df['Collections Status'] == by_product]['Principal Balance'].sum()
            D = loans_df[loans_df['Collections Status'] == by_product].shape[0] - beg_of_month_loans_df[beg_of_month_loans_df['Collections Status'] == by_product].shape[0]
            E = loans_df[loans_df['Collections Status'] == by_product]['Principal Balance'].sum() - beg_of_month_loans_df[beg_of_month_loans_df['Collections Status'] == by_product]['Principal Balance'].sum()
            F = loans_df[loans_df['Collections Status'] == by_product].shape[0] - beg_of_year_loans_df[beg_of_year_loans_df['Collections Status'] == by_product].shape[0]
            G = loans_df[loans_df['Collections Status'] == by_product]['Principal Balance'].sum() - beg_of_year_loans_df[beg_of_year_loans_df['Collections Status'] == by_product]['Principal Balance'].sum()
        write_entry_summary(summary_sheet, A, B, C, D, E, F, G, '', '', by_product_count)
        summary_sheet['D' + str(by_product_count)].number_format = '#,##0'
        summary_sheet['E' + str(by_product_count)].number_format = '#,##0'
        summary_sheet['F' + str(by_product_count)].number_format = '#,##0'
        summary_sheet['G' + str(by_product_count)].number_format = '#,##0'
        by_product_count += 1
    set_border(summary_sheet, 'A' + str(by_product_border_count) + ':' + 'G' + str(by_product_count - 1))
    summary_sheet['A' + str(by_product_count)] = 'Total Book Size'
    summary_sheet['A' + str(by_product_count)].font = Font(size=10, bold=True)
    summary_sheet['A' + str(by_product_count)].number_format = '@'
    summary_sheet['B' + str(by_product_count)].font = Font(size=10, bold=True)
    summary_sheet['B' + str(by_product_count)].value = loans_df[loans_df['Products'] == 'CSL'].shape[0]
    summary_sheet['B' + str(by_product_count)].number_format = '#,##0'
    summary_sheet['C' + str(by_product_count)].font = Font(size=10, bold=True)
    summary_sheet['C' + str(by_product_count)].value = loans_df[loans_df['Products'] == 'CSL']['Principal Balance'].sum()
    summary_sheet['C' + str(by_product_count)].number_format = '#,##0'
    summary_sheet['D' + str(by_product_count)].font = Font(size=10, bold=True)
    summary_sheet['D' + str(by_product_count)].value = loans_df[loans_df['Products'] == 'CSL'].shape[0] - beg_of_month_loans_df[beg_of_month_loans_df['Products'] == 'CSL'].shape[0]
    summary_sheet['D' + str(by_product_count)].number_format = '#,##0'
    summary_sheet['E' + str(by_product_count)].font = Font(size=10, bold=True)
    summary_sheet['E' + str(by_product_count)].value = loans_df[loans_df['Products'] == 'CSL']['Principal Balance'].sum() - beg_of_month_loans_df[beg_of_month_loans_df['Products'] == 'CSL']['Principal Balance'].sum()
    summary_sheet['E' + str(by_product_count)].number_format = '#,##0'
    summary_sheet['F' + str(by_product_count)].font = Font(size=10, bold=True)
    summary_sheet['F' + str(by_product_count)].value = loans_df[loans_df['Products'] == 'CSL'].shape[0] - beg_of_year_loans_df[beg_of_year_loans_df['Products'] == 'CSL'].shape[0]
    summary_sheet['F' + str(by_product_count)].number_format = '#,##0'
    summary_sheet['G' + str(by_product_count)].font = Font(size=10, bold=True)
    summary_sheet['G' + str(by_product_count)].value = loans_df[loans_df['Products'] == 'CSL']['Principal Balance'].sum() - beg_of_year_loans_df[beg_of_year_loans_df['Products'] == 'CSL']['Principal Balance'].sum()
    summary_sheet['G' + str(by_product_count)].number_format = '#,##0'
    by_product_count += 1
    collection_count = 2
    print('Work Through Collection')
    for collection_officer in tqdm(collection_officers, colour='green'):
        loan_book_col_filter_df = loans_df[(loans_df['Collection_Officer'] == collection_officer) & (loans_df['Products'] == 'CSL')]
        start_of_year_filter_df = beg_of_year_loans_df[(beg_of_year_loans_df['Collection_Officer'] == collection_officer) & (beg_of_year_loans_df['Products'] == 'CSL')]
        start_of_month_filter_df = beg_of_month_loans_df[(beg_of_month_loans_df['Collection_Officer'] == collection_officer) & (beg_of_month_loans_df['Products'] == 'CSL')]
        yesterdays_filter_df = yesterdays_loans_df[(yesterdays_loans_df['Collection_Officer'] == collection_officer) & (yesterdays_loans_df['Products'] == 'CSL')]
        work_through_collections(loan_book_col_filter_df, start_of_year_filter_df, start_of_month_filter_df, yesterdays_filter_df, collection_officer, collection_count, collections_sheet)
        collection_count += 1
    set_border(collections_sheet, 'A2:O' + str(collection_count - 1))
    branch_summary_count = 2
    print('Work Through Branch Summary')
    for branch in tqdm(cs_branch_list, colour='green'):
        loan_book_filter_df = loans_df[(loans_df['Branch'] == branch) & (loans_df['Products'] == 'CSL')]
        start_of_year_filter_df = beg_of_year_loans_df[(beg_of_year_loans_df['Branch'] == branch) & (beg_of_year_loans_df['Products'] == 'CSL')]
        start_of_month_filter_df = beg_of_month_loans_df[(beg_of_month_loans_df['Branch'] == branch) & (beg_of_month_loans_df['Products'] == 'CSL')]
        yesterdays_filter_df = yesterdays_loans_df[(yesterdays_loans_df['Branch'] == branch) & (yesterdays_loans_df['Products'] == 'CSL')]
        work_through_cs_branch_summary(loan_book_filter_df, start_of_year_filter_df, start_of_month_filter_df, yesterdays_filter_df, branch, branch_summary_count, cs_branch_summary_sheet)
        branch_summary_count += 1
    set_border(cs_branch_summary_sheet, 'A2:M' + str(branch_summary_count - 1))
    ro_count = 2
    print('Work Through Relationship Managers')
    for ro in tqdm(relationship_managers, colour='green'):
        loan_book_ro_filter_df = loans_df[loans_df['Relationship/Recoveries Officers'] == ro]
        start_of_year_filter_df = beg_of_year_loans_df[beg_of_year_loans_df['Relationship/Recoveries Officers'] == ro]
        start_of_month_filter_df = beg_of_month_loans_df[beg_of_month_loans_df['Relationship/Recoveries Officers'] == ro]
        yesterdays_filter_df = yesterdays_loans_df[yesterdays_loans_df['Relationship/Recoveries Officers'] == ro]
        work_through_lbf(loan_book_ro_filter_df, start_of_year_filter_df, start_of_month_filter_df, yesterdays_filter_df, ro, ro_count, ro_sheet)
        ro_count += 1
    set_border(ro_sheet, 'A2:N' + str(ro_count - 1))
    lbf_branch_count = 2
    print('Work Through LBF Branches')
    for lbf_branch in tqdm(lbf_branches, colour='green'):
        loan_book_branch_filter_df = loans_df[loans_df['Branch'] == lbf_branch]
        start_of_year_filter_df = beg_of_year_loans_df[beg_of_year_loans_df['Branch'] == lbf_branch]
        start_of_month_filter_df = beg_of_month_loans_df[beg_of_month_loans_df['Branch'] == lbf_branch]
        yesterdays_filter_df = yesterdays_loans_df[yesterdays_loans_df['Branch'] == lbf_branch]
        work_through_lbf_branch(loan_book_branch_filter_df, start_of_year_filter_df, start_of_month_filter_df, yesterdays_filter_df, lbf_branch, lbf_branch_count, lbf_branch_sheet)
        lbf_branch_count += 1
    set_border(lbf_branch_sheet, 'A2:N' + str(lbf_branch_count - 1))
    private_sector_count = 2
    print('Work Through Private Sector Employers')
    for employer in tqdm(private_sector_employers, colour='green'):
        loan_book_branch_filter_df = loans_df[loans_df['Employer_Name (Client)'] == employer]
        start_of_year_filter_df = beg_of_year_loans_df[beg_of_year_loans_df['Employer_Name (Client)'] == employer]
        start_of_month_filter_df = beg_of_month_loans_df[beg_of_month_loans_df['Employer_Name (Client)'] == employer]
        yesterdays_filter_df = yesterdays_loans_df[yesterdays_loans_df['Employer_Name (Client)'] == employer]
        work_through_private_sector_employer(loan_book_branch_filter_df, start_of_year_filter_df, start_of_month_filter_df, yesterdays_filter_df, employer, private_sector_count, private_sector_sheet)
        private_sector_count += 1
    set_border(private_sector_sheet, 'A2:H' + str(private_sector_count - 1))
    sme_branch_count = 2
    print('Work Through SME Branches')
    for branch in tqdm(sme_branches, colour='green'):
        loan_book_filter_df = loans_df[loans_df['Branch'] == branch]
        start_of_year_filter_df = beg_of_year_loans_df[beg_of_year_loans_df['Branch'] == branch]
        start_of_month_filter_df = beg_of_month_loans_df[beg_of_month_loans_df['Branch'] == branch]
        yesterdays_filter_df = yesterdays_loans_df[yesterdays_loans_df['Branch'] == branch]
        work_through_sme(loan_book_filter_df, start_of_year_filter_df, start_of_month_filter_df, yesterdays_filter_df, branch, sme_branch_count, sme_sheet)
        sme_branch_count += 1
    set_border(sme_sheet, 'A2:O' + str(sme_branch_count - 1))
    sme_agent_count = 2
    print('Work Through SME Agents')
    for lo in tqdm(sme_loan_officers, colour='green'):
        loan_book_filter_df = loans_df[loans_df['Sales Representative (Loan)'] == lo]
        start_of_year_filter_df = beg_of_year_loans_df[beg_of_year_loans_df['Sales Representative (Loan)'] == lo]
        start_of_month_filter_df = beg_of_month_loans_df[beg_of_month_loans_df['Sales Representative (Loan)'] == lo]
        yesterdays_filter_df = yesterdays_loans_df[yesterdays_loans_df['Sales Representative (Loan)'] == lo]
        work_through_sme_agent(loan_book_filter_df, start_of_year_filter_df, start_of_month_filter_df, yesterdays_filter_df, lo, sme_agent_count, sme_agent_sheet)
        sme_agent_count += 1
    set_border(sme_agent_sheet, 'A2:O' + str(sme_agent_count - 1))
    maziwa_branch_count = 2
    print('Work Through Maziwa Branches')
    for branch in tqdm(maziwa_branches, colour='green'):
        loan_book_filter_df = loans_df[loans_df['Branch'] == branch]
        start_of_year_filter_df = beg_of_year_loans_df[beg_of_year_loans_df['Branch'] == branch]
        start_of_month_filter_df = beg_of_month_loans_df[beg_of_month_loans_df['Branch'] == branch]
        yesterdays_filter_df = yesterdays_loans_df[yesterdays_loans_df['Branch'] == branch]
        work_through_maziwa(loan_book_filter_df, start_of_year_filter_df, start_of_month_filter_df, yesterdays_filter_df, branch, maziwa_branch_count, maziwa_sheet)
        maziwa_branch_count += 1
    set_border(maziwa_sheet, 'A2:O' + str(maziwa_branch_count - 1))
    maziwa_agent_count = 2
    print('Work Through Maziwa Agents')
    for lo in tqdm(maziwa_loan_officers, colour='green'):
        loan_book_filter_df = loans_df[loans_df['Sales Representative (Loan)'] == lo]
        start_of_year_filter_df = beg_of_year_loans_df[beg_of_year_loans_df['Sales Representative (Loan)'] == lo]
        start_of_month_filter_df = beg_of_month_loans_df[beg_of_month_loans_df['Sales Representative (Loan)'] == lo]
        yesterdays_filter_df = yesterdays_loans_df[yesterdays_loans_df['Sales Representative (Loan)'] == lo]
        work_through_maziwa_agent(loan_book_filter_df, start_of_year_filter_df, start_of_month_filter_df, yesterdays_filter_df, lo, maziwa_agent_count, maziwa_agent_sheet)
        maziwa_agent_count += 1
    set_border(maziwa_agent_sheet, 'A2:O' + str(maziwa_agent_count - 1))
    easy_phone_count = 2
    print('Work Through Easy Phone Sales Agents')
    for agent in tqdm(easy_phone_agents, colour='green'):
        loan_book_agent_filter_df = loans_df[(loans_df['Sales Representative (Loan)'] == agent) & (loans_df['Products'] == 'Eezy-Phone')]
        start_of_year_filter_df = beg_of_year_loans_df[(beg_of_year_loans_df['Sales Representative (Loan)'] == agent) & (beg_of_year_loans_df['Products'] == 'Eezy-Phone')]
        start_of_month_filter_df = beg_of_month_loans_df[(beg_of_month_loans_df['Sales Representative (Loan)'] == agent) & (beg_of_month_loans_df['Products'] == 'Eezy-Phone')]
        yesterdays_filter_df = yesterdays_loans_df[(yesterdays_loans_df['Sales Representative (Loan)'] == agent) & (yesterdays_loans_df['Products'] == 'Eezy-Phone')]
        work_through_easy_phone(loan_book_agent_filter_df, start_of_year_filter_df, start_of_month_filter_df, yesterdays_filter_df, agent, easy_phone_count, easy_phone_sheet)
        easy_phone_count += 1
    set_border(easy_phone_sheet, 'A2:O' + str(easy_phone_count - 1))
    working_capital_count = 2
    print('Work Through Working Capital Agents')
    for agent in tqdm(working_capital_agents, colour='green'):
        loan_book_agent_filter_df = loans_df[(loans_df['Sales Representative (Loan)'] == agent) & (loans_df['Products'] == 'Working Capital')]
        start_of_year_filter_df = beg_of_year_loans_df[(beg_of_year_loans_df['Sales Representative (Loan)'] == agent) & (beg_of_year_loans_df['Products'] == 'Working Capital')]
        start_of_month_filter_df = beg_of_month_loans_df[(beg_of_month_loans_df['Sales Representative (Loan)'] == agent) & (beg_of_month_loans_df['Products'] == 'Working Capital')]
        yesterdays_filter_df = yesterdays_loans_df[(yesterdays_loans_df['Sales Representative (Loan)'] == agent) & (yesterdays_loans_df['Products'] == 'Working Capital')]
        work_through_working_capital(loan_book_agent_filter_df, start_of_year_filter_df, start_of_month_filter_df, yesterdays_filter_df, agent, working_capital_count, working_capital_sheet)
        working_capital_count += 1
    set_border(working_capital_sheet, 'A2:O' + str(working_capital_count - 1))
    for worksheet in output_book.worksheets:
        for column_cells in worksheet.columns:
            width = max((len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells if cell.coordinate not in worksheet.merged_cells), default=0)
            worksheet.column_dimensions[get_column_letter(column_cells[0].column)].width = (width + 2) * 1.2
    output_book.save('collections_report_{}.xlsx'.format(day_month_year))
if __name__ == '__main__':
    try:
        main()
        print('\nDone. Report saved as collections_report_{}.xlsx'.format(day_month_year))
    except Exception:
        import traceback
        traceback.print_exc()
    input('\nPress Enter to close...')
