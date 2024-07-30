# Import Python libraries.
from num2words import num2words
from tqdm import tqdm
import warnings
import argparse
import pdfkit
import PyPDF2
import pandas as pd
import os
import logging
import sys
from datetime import datetime
from pathlib import Path



filename = sys.argv[1]  # Access the first argument
lut = sys.argv[2]  # Access the second argument

 
# Initialising and Defining LOGGER object.
log_file_path = os.path.join(os.path.join(os.path.dirname(__file__)), "log")
# Create log file path if not exists.
Path(log_file_path).mkdir(parents=True, exist_ok=True)
# Initialize logger object
log = logging.getLogger(__name__)
# Define format of logging.
formatter = logging.Formatter(
    '%(asctime)s - %(name)-12s - %(levelname)-4s - %(filename)s - %(funcName)s -%(lineno)d - %(message)s')
# Define logging file path.
file_handle = logging.FileHandler(
    log_file_path+f'/{datetime.now().strftime("%d_%m_%Y_%H_%M_%S")}.log')
file_handle.setFormatter(formatter)
# Adding handler information to logger.
log.addHandler(file_handle)
log.setLevel(logging.INFO)


# Import python Libraries.
log.info("Importing python libraries.")

warnings.filterwarnings("ignore")

log.info("Initialising and configuring argument parser.")
# Initialising and configuring argument parser.
parser = argparse.ArgumentParser(
    prog='Invoice Generation',
    description='The program is to automate invoice generation using Pandas, pdfkit and other relavent libraries in Python Programming language',
    epilog='Please contact the relavent teams for further inquire')

log.info("Initialising argparser arguments.")
# Initilaising arguments.
parser.add_argument('-f', '--filename', type=str,
                    help="Please provide the excel filename with invoices information uploaded at upload folder.")
parser.add_argument('-l', '--lut', type=str,
                    help="Please provide the lut pdf file name uploaded in upload folder")
args = parser.parse_args()

log.info("Defining constant variables includes HTML Template, file and folder paths.")
# Constant variables.
html_header = """
				<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
				<html>
					<head>
						<meta http-equiv="content-type" content="text/html; charset=iso-8859-1"/>
				<meta name="viewport" content="width=device-width, initial-scale=1" />
						<title></title>
			  """
html_styling = """
				<style type="text/css">
				body,div,table,thead,tbody,tfoot,tr,th,td,p { font-family:"Arial"; font-size:x-small }
				a.comment-indicator:hover + comment { background:#ffd; position:absolute; display:block; border:1px solid black; padding:0.5em;  } 
				a.comment-indicator { background:red; display:inline-block; border:1px solid black; width:0.5em; height:0.5em;  } 
				comment { display:none;  } 
				</style>
			   """
html_body = """
				</head>
				<body>
					<table cellspacing="0" border="0" >
						<colgroup width="135"></colgroup>
						<colgroup width="206"></colgroup>
						<colgroup width="98"></colgroup>
						<colgroup width="86"></colgroup>
						<colgroup width="147"></colgroup>
						<colgroup width="81"></colgroup>
						<colgroup width="63"></colgroup>
						<colgroup width="67"></colgroup>
						<tr>
							<td colspan=2 rowspan=6 height="80" align="left" valign=bottom><font
									face="Times New Roman"><br><img
										src="https://otsi-global.com/wp-content/uploads/2022/02/Final-Logo-Colour-and-Proportion-2048x837.png" width=200 height=80>
								</font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
						</tr>
						<tr>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
						</tr>
						<tr>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><b><font color="#06456B"><br></font></b></td>
							<td colspan=3 align="right" valign=bottom><b><font color="#06456B">Object
										Technology Solutions</font></b></td>
						</tr>
						<tr>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><b><font color="#06456B"><br></font></b></td>
							<td colspan=3 align="right" valign=bottom><b><font color="#06456B">India
										Private Limited</font></b></td>
						</tr>
						<tr>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
						</tr>
						<tr>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
						</tr>
						<tr>
							<td height="20" align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td colspan=4 align="center"
								valign=bottom><i><font face="Times New Roman" size=4>(Original for Buyer)</font></i></td>
						</tr>
						<tr>
							<td style="border-top:2px solid #000000;border-right:2px solid
								#000000;border-left:2px solid #000000" colspan=8
								height="29" align="center" valign=middle><b><u><font face="Times New Roman"
											size=4>TAX INVOICE</font></u></b></td>
						</tr>
						<tr>
							<td style="border:2px solid #000000" colspan=8
								height="48" align="center" valign=middle><font face="Times New Roman"
									size=3>(Supply Meant For Export/Supply to SEZ Unit or SEZ Developer For
									Authorised Operations Under Bond or Letter Of Undertaking Without Payment
									of IGST)</font></td>
						</tr>
						<tr>
							<td style=" border-left: 2px solid #000000;"
								rowspan=3 height="30" align="left"
								valign=middle><b><font face="Times New Roman">Vendor Address :</font></b></td>
							<td style="border-left: 2px solid #000000;"
								colspan=3 rowspan=3 align="left"
								valign=top><font face="Times New Roman">Object Technology Solutions India
									Pvt. Ltd Phoenix Info city, SEZ ,Hitech City - 2, Hyderabad &ndash;
									500081,Telangana, India</font></td>
							<td style="border-left: 2px solid #000000;"
								rowspan=2
								align="left" valign=bottom><font face="Times New Roman"><b>Invoice Number :</b></font></td>
							<td style="border-left: 2px solid #000000; border-right: 2px solid #000000"
								colspan=3
								rowspan=2 align="left" valign=bottom><font face="Times New Roman">{inv_no}</font></td>
						</tr>
						<tr>
						</tr>
						<tr>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000;"
								align="left" valign=bottom><font face="Times New Roman"><b>Invoice Period:</b></font></td>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000; border-right: 2px solid #000000" colspan=3
								align="left" valign=bottom sdnum="1033;1033;MMM-YY"><font face="Times New
									Roman">{inv_prd}</font></td>
						</tr>
						<tr>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000" height="20" align="left" valign=bottom><b><font
										face="Times New Roman">GSTIN/UN :</font></b></td>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000;" colspan=3
								align="left" valign=bottom><font face="Times New Roman">36AABCO3022C1Z4</font></td>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000;"
								align="left" valign=bottom><font face="Times New Roman"><b>Invoice Date :</b></font></td>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000; border-right: 2px solid #000000" colspan=3
								align="left" valign=bottom sdval="45055" sdnum="1033;1033;M/D/YYYY"><font
									face="Times New Roman">{inv_dt}</font></td>
						</tr>
						<tr>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000" height="20" align="left" valign=bottom><b><font
										face="Times New Roman">CIN :</font></b></td>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000;" colspan=3
								align="left" valign=bottom><font face="Times New Roman">U72200TS2010PTC068549</font></td>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000;"
								align="left" valign=middle><font face="Times New Roman"><b>WO Number</b></font></td>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000; border-right: 2px solid #000000" colspan=3
								align="left" valign=middle sdnum="1033;1033;M/D/YYYY"><font face="Times New
									Roman">{wonum}</font></td>
						</tr>
						<tr>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000" height="20" align="left" valign=bottom><b><font
										face="Times New Roman">Company's PAN :</font></b></td>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000;" colspan=3
								align="left" valign=bottom><font face="Times New Roman">AABCO3022C</font></td>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000; "
								align="left" valign=middle><font face="Times New Roman"><b>WO Date</b></font></td>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000; border-right: 2px solid #000000" colspan=3
								align="left" valign=middle sdval="44963" sdnum="1033;1033;M/D/YYYY"><font
									face="Times New Roman">{wodt}</font></td>
						</tr>
						<tr>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000" rowspan=3 height="30" align="left"
								valign=middle><b><font face="Times New Roman">LUT Details: </font></b></td>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000;" colspan=3
								rowspan=3 align="left" valign=middle><font face="Times New Roman">ARN No:
									AD3601230077965 dated 16.01.2023 valid from 01.04.2023 to 31.03.2024</font></td>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000; "
								align="left" valign=middle><font face="Times New Roman"><b>Consolidated
									Invoice ID</b></font></td>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000; border-right: 2px solid #000000" colspan=3
								align="left" valign=middle><font face="Times New Roman">{ciid}</font></td>
						</tr>
						<tr>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000; " rowspan=2
								align="left" valign=bottom><font face="Times New Roman"><b>Reverse Charge</b></font></td>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000; border-right: 2px solid #000000" colspan=3
								rowspan=2 align="left" valign=bottom><font face="Times New Roman">NO</font></td>
						</tr>
						<tr>
						</tr>
						<tr>
							<td style="border-top: 1px solid #000000; border-bottom: 1px solid #000000;
								border-left: 2px solid #000000;"
								height="23" align="left" valign=middle><b><font face="Times New Roman">E-Mail
										:</font></b></td>
							<td style="border-top: 1px solid #000000; border-bottom: 1px solid #000000;
								border-left: 2px solid #000000;" colspan=3
								align="left" valign=middle><u><font color="#0000FF"><a
											href="mailto:venkat@otsi-usa.com">venkat@otsi-usa.com</a></font></u></td>
							<td style="border-top: 1px solid #000000; border-bottom: 1px solid #000000;
								border-left: 2px solid #000000; "
								align="left" valign=bottom><font face="Times New Roman"><b>Place Of Supply:</b></font></td>
							<td style="border-top: 1px solid #000000; border-bottom: 1px solid #000000;
								border-left: 2px solid #000000; border-right: 2px solid #000000" colspan=3
								align="left" valign=bottom><font face="Times New Roman">{pos}</font></td>
						</tr>
						<tr style="border: 2px solid #000000 !important;">
							<td style="border-top: 0px solid #000000; border-bottom: 1px solid #000000;
								border-left: 2px solid #000000; border-right: 2px  #000000"
								height="40" align="left" valign=middle><b><font face="Times New Roman">Ship
										To Address :</font></b></td>
							<td style="border-top: 1px solid #000000; border-bottom: 1px solid #000000;
								border-left: 2px solid #000000; border-right: 2px solid #000000" colspan=7
								align="left" valign=top><b><font face="Times New Roman">{sta}</font></b></td>
						</tr>
						<tr>
							<td style="border-top: 1px solid #000000; border-bottom: 1px solid #000000;
								border-left: 2px solid #000000; border-right: 2px  #000000"
								height="18" align="left" valign=bottom><b><font face="Times New Roman">GSTIN/UN
										:</font></b></td>
							<td style="border-top: 1px solid #000000; border-bottom: 1px solid #000000;
								border-left: 2px solid #000000; border-right: 2px solid #000000" colspan=7
								align="left" valign=top><b><font face="Times New Roman">{gstin}</font></b></td>
						</tr>
						<tr>
							<td style="border-top: 1px solid #000000; border-bottom: 1px solid #000000;
								border-left: 2px solid #000000; border-right: 2px  #000000"
								height="40" align="left" valign=middle><b><font face="Times New Roman">Bill
										To Address:</font></b></td>
							<td style="border-top: 1px solid #000000; border-bottom: 1px solid #000000;
								border-left: 2px solid #000000; border-right: 2px solid #000000" colspan=7
								align="left" valign=top><b><font face="Times New Roman">{bta}</font></b></td>
						</tr>
						<tr>
							<td style="border-top: 1px solid #000000; border-bottom: 1px solid #000000;
								border-left: 2px solid #000000; border-right: 2px  #000000"
								height="18" align="left" valign=bottom><b><font face="Times New Roman">GSTIN/UN
										:</font></b></td>
							<td style="border-top: 1px solid #000000; border-bottom: 1px solid #000000;
								border-left: 2px solid #000000; border-right: 2px solid #000000" colspan=7
								align="left" valign=top><b><font face="Times New Roman">{bgstin}</font></b></td>
						</tr>
						<tr>
							<td style="border-top: 1px solid #000000; border-bottom: 1px solid #000000;
								border-left: 2px solid #000000; border-right: 2px #000000"
								height="30" align="center" valign=bottom><b><font face="Times New Roman">Sl
										No</font></b></td>
							<td style="border-top: 1px solid #000000; border-bottom: 1px solid #000000;
								border-left: 2px solid #000000; border-right: 2px  #000000"
								align="left" valign=bottom><b><font face="Times New Roman">Item Description</font></b></td>
							<td style="border-top: 1px solid #000000; border-bottom: 1px solid #000000;
								border-left: 2px solid #000000; border-right: 2px  #000000"
								align="center" valign=bottom><b><font face="Times New Roman">Is the item a
										GOOD (G) or SERVICE (S) </font></b></td>
							<td style="border-top: 1px solid #000000; border-bottom: 1px solid #000000;
								border-left: 2px solid #000000; border-right: 2px #000000"
								align="center" valign=bottom><b><font face="Times New Roman">HSN or SAC
										code </font></b></td>
							<td style="border-top: 1px solid #000000; border-bottom: 1px solid #000000;
								border-left: 2px solid #000000; border-right: 2px  #000000"
								align="center" valign=bottom><b><font face="Times New Roman">Quantity</font></b></td>
							<td style="border-top: 1px solid #000000; border-bottom: 1px solid #000000;
								border-left: 2px solid #000000; border-right: 2px  #000000"
								align="right" valign=bottom><b><font face="Times New Roman">Item Price
									</font></b></td>
							<td style="border-top: 1px solid #000000; border-bottom: 1px solid #000000;
								border-left: 2px solid #000000; border-right: 2px solid #000000" colspan=2
								align="center" valign=bottom><b><font face="Times New Roman">Gross Amount
									</font></b></td>
						</tr>
						<tr>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000;" rowspan=2
								height="40" align="center" valign=middle sdval="1" sdnum="1033;"><font
									face="Times New Roman">1</font></td>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000;" rowspan=2
								align="left" valign=middle><font face="American Typewriter">{wrkr_name}
									{itmdisc}</font></td>
							<td style="border-top: 1px solid #000000;
								border-left: 1px solid #000000; " rowspan=2
								align="center" valign=middle><font face="American Typewriter">S</font></td>
							<td style="border-top: 1px solid #000000;
								border-left: 1px solid #000000; " rowspan=2
								align="center" valign=middle sdval="998513" sdnum="1033;"><font
									face="American Typewriter">998513</font></td>
							<td style="border-top: 1px solid #000000;
								border-left: 1px solid #000000; " rowspan=2
								align="center" valign=middle sdval="{bhr}" sdnum="1033;"><font
									face="American
									Typewriter">{bhr}</font></td>
							<td style="border-top: 1px solid #000000;
								border-left: 1px solid #000000;" rowspan=2
								align="center" valign=middle sdval="{brt}" sdnum="1033;0;_
								[$&#8377;-4009]
								* #,##0_ ;_ [$&#8377;-4009] * -#,##0_ ;_ [$&#8377;-4009] * &quot;-&quot;??_
								;_ @_ "><font face="Times New Roman"> &#8377; {brt} </font></td>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000; border-right: 2px solid #000000" colspan=2
								rowspan=2 align="center" valign=middle sdval="{silia}" sdnum="1033;0;_
								[$&#8377;-4009] * #,##0_ ;_ [$&#8377;-4009] * -#,##0_ ;_ [$&#8377;-4009] *
								&quot;-&quot;??_ ;_ @_ "><font face="Times New Roman"> &#8377; {silia}
								</font></td>
						</tr>
						<tr>8513
						</tr>
						<tr>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000; "
								height="17" align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000;" colspan=2
								align="center" valign=bottom><font face="Times New Roman"><br></font></td>
							<td style="border-top: 1px solid #000000;
								border-left: 1px solid #000000;"
								align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td style="border-top: 1px solid #000000;
								border-left: 1px solid #000000;" colspan=2
								align="center" valign=bottom><b><font face="Times New Roman">Gross Total</font></b></td>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000; border-right: 2px solid #000000" colspan=2
								align="center" valign=bottom sdval="{silia}" sdnum="1033;0;_
								[$&#8377;-445] * #,##0_ ;_ [$&#8377;-445] * -#,##0_ ;_ [$&#8377;-445] *
								&quot;-&quot;??_ ;_ @_ "><b><font face="Times New Roman"> &#8377; {silia}
									</font></b></td>
						</tr>
						<tr>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000; "
								height="17" align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000" align="center" valign=bottom><font
									face="Times New Roman"><br></font></td>
							<td style="border-top: 1px solid #000000;
								" align="center" valign=bottom><font
									face="Times New Roman"><br></font></td>
							<td style="border-top: 1px solid #000000;
								border-left: 1px solid #000000; "
								align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td style="border-top: 1px solid #000000;
								border-left: 1px solid #000000; " colspan=2
								align="center" valign=bottom><b><font face="Times New Roman">Less : MSP<br>Fee-2.6
										%</font></b></td>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000; border-right: 2px solid #000000" colspan=2
								align="center" valign=bottom sdval="{slmspfee}" sdnum="1033;0;_
								[$&#8377;-445] * #,##0_ ;_ [$&#8377;-445] * -#,##0_ ;_ [$&#8377;-445] *
								&quot;-&quot;??_ ;_ @_ "><b><font face="Times New Roman"> &#8377;
										{slmspfee}
									</font></b></td>
						</tr>
						<tr>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000;"
								height="17" align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000" align="center" valign=bottom><font
									face="Times New Roman"><br></font></td>
							<td style="border-top: 1px solid #000000;" align="center" valign=bottom><font
									face="Times New Roman"><br></font></td>
							<td style="border-top: 1px solid #000000;
								border-left: 1px solid #000000;"
								align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td style="border-top: 1px solid #000000;
								border-left: 1px solid #000000;" colspan=2
								align="center" valign=bottom><b><font face="Times New Roman">Taxable Value</font></b></td>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000; border-right: 2px solid #000000" colspan=2
								align="center" valign=bottom sdval="{stiawgst}" sdnum="1033;0;_
								[$&#8377;-445] * #,##0_ ;_ [$&#8377;-445] * -#,##0_ ;_ [$&#8377;-445] *
								&quot;-&quot;??_ ;_ @_ "><b><font face="Times New Roman"> &#8377;
										{stiawgst}
									</font></b></td>
						</tr>
						<tr>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000;"
								height="17" align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000;" colspan=2
								align="center" valign=bottom><font face="Times New Roman"><br></font></td>
							<td style="border-top: 1px solid #000000;
								border-left: 1px solid #000000;"
								align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td style="border-top: 1px solid #000000;
								border-left: 1px solid #000000;" colspan=2
								align="center" valign=bottom><b><font face="Times New Roman">IGST </font></b></td>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000; border-right: 2px solid #000000" colspan=2
								align="center" valign=bottom sdval="{gst_value}" sdnum="1033;0;_
								[$&#8377;-445]
								*
								#,##0_ ;_ [$&#8377;-445] * -#,##0_ ;_ [$&#8377;-445] * &quot;-&quot;??_ ;_
								@_ "><b><font face="Times New Roman"> &#8377;{gst_value}</font></b></td>
						</tr>
						<tr>
							<td style="border-top: 1px solid #000000; border-bottom: 1px solid #000000;
								border-left: 2px solid #000000;border-right: px solid #000000;"
								height="18" align="left" valign=bottom><font face="Times New Roman"><br></font></td>
								<td style="border-top: 1px solid #000000;border-bottom: 1px solid #000000;
								border-left: 2px solid #000000;border-right: 1px  #000000;" colspan=2
								align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td style="border-top: 1px solid #000000;border-left: 1px solid
								#000000;border-bottom: 1px solid #000000;border-right: 1px  #000000;" align="left" valign=bottom><font
									face="Times New Roman"><br></font></td>
							<td style="border-top: 1px solid #000000; border-bottom: 1px solid #000000;
								border-left: 1px solid #000000;" colspan=2
								align="center" valign=bottom><b><font face="Times New Roman">Total Amount</font></b></td>
							<td style="border-top: 2px solid #000000; border-left: 2px solid #000000;
								border-right: 2px solid #000000;border-bottom: 1px solid #000000;"
								colspan=2 align="left" valign=bottom
								sdval="{sumta}" sdnum="1033;0;_ [$&#8377;-445] * #,##0_ ;_
								[$&#8377;-445] * -#,##0_ ;_ [$&#8377;-445] * &quot;-&quot;??_ ;_ @_ "><b><font
										face="Times New Roman"> &#8377; {sumta} </font></b></td>
						</tr>
						<tr>
							<td style="border-top: 1px solid #000000; border-bottom: 1px solid #000000;
								border-left: 2px solid #000000; border-right: 1px  #000000"
								height="18" align="left" valign=bottom><b><font face="Times New Roman">Amount
										in Words :</font></b></td>
							<td style="border-top: 1px solid #000000; border-bottom: 1px solid #000000;
								border-left: 2px solid #000000; border-right: 2px solid #000000" colspan=7
								align="left" valign=bottom><font face="Times New Roman">{sumttotal_words}</font></td>
						</tr>
						<tr>
							<td style=" border-left: 2px solid #000000;border-top: 1px solid #000000;border-bottom: 1px solid #000000;"
								height="17" align="left" valign=bottom><b><font
										face="Times New Roman">HSN/SAC</font></b></td>
							<td style=" border-left: 2px solid #000000;border-top: 1px solid #000000;border-bottom: 1px solid #000000;"
								align="left" valign=middle><b><font
										face="Times New Roman">Taxable Value</font></b></td>
							<td style=" border-left: 2px solid #000000;border-top: 1px solid #000000;border-bottom: 1px solid #000000;"
								colspan=2 align="center" valign=middle><b><font
										face="Times New Roman">CGST</font></b></td>
							<td style=" border-left: 2px solid #000000;border-top: 1px solid #000000;border-bottom: 1px solid #000000;"
								colspan=2 align="center" valign=bottom><b><font
										face="Times New Roman">SGST</font></b></td>
							<td style=" border-left: 2px solid #000000;border-top: 1px solid
								#000000;border-right: 2px solid #000000;border-bottom: 1px solid #000000;" colspan=2 align="center"
								valign=middle><b><font
										face="Times New Roman">IGST</font></b></td>
						</tr>
						<tr>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000;"
								height="17" align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000;"
								align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000;border-bottom: 1px solid #000000;"
								align="left" valign=bottom><b><font face="Times New Roman">Rate</font></b></td>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000;border-bottom: 1px solid #000000;"
								align="left" valign=bottom><b><font face="Times New Roman">Amount</font></b></td>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000;border-bottom: 1px solid #000000;"
								align="left" valign=bottom><b><font face="Times New Roman">Rate</font></b></td>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000;border-bottom: 1px solid #000000;"
								align="left" valign=bottom><b><font face="Times New Roman">Amount</font></b></td>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000;border-bottom: 1px solid #000000;"
								align="left" valign=bottom><b><font face="Times New Roman">Rate</font></b></td>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000; border-right: 2px solid #000000;border-bottom: 1px solid #000000;"
								align="left" valign=bottom><b><font face="Times New Roman">Amount</font></b></td>
						</tr>
						<tr>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000;"
								height="17" align="left" valign=bottom sdval="998513" sdnum="1033;"><font
									face="Times New Roman">998513</font></td>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000;"
								align="right" valign=bottom sdval="{stiawgst}" sdnum="1033;0;0"><font
									face="Times New Roman">{stiawgst}</font></td>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000;"
								align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000;"
								align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000;"
								align="left" valign=bottom sdnum="1033;0;0%"><font face="Times New Roman"><br></font></td>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000;"
								align="left" valign=bottom sdnum="1033;0;0"><font face="Times New Roman"><br></font></td>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000;"
								align="right" valign=bottom sdval="{gst}" sdnum="1033;0;0%"><font
									face="Times
									New Roman">{gst}%</font></td>
							<td style="border-top: 1px solid #000000;
								border-left: 2px solid #000000; border-right: 2px solid #000000"
								align="right" valign=bottom sdval="{gst_value}" sdnum="1033;0;0"><font
									face="Times
									New Roman">{gst_value}</font></td>
						</tr>
						<tr>
							<td style="border-top: 1px solid #000000; border-bottom: 1px solid #000000;
								border-left: 2px solid #000000" height="18" align="center" valign=middle><b><font
										face="Times New Roman">Total</font></b></td>
							<td style="border-top: 1px solid #000000; border-left: 2px solid #000000;
								border-bottom: 1px solid #000000" align="right" valign=middle
								sdnum="1033;0;_(* #,##0.00_);_(* \(#,##0.00\);_(* &quot;-&quot;??_);_(@_)"><b><font
										face="Times New Roman"><br></font></b></td>
							<td style="border-top: 1px solid #000000; border-left: 2px solid #000000;
								border-bottom: 1px solid #000000" align="left" valign=bottom><font
									face="Times New Roman"><br></font></td>
							<td style="border-top: 1px solid #000000; border-left: 2px solid #000000;
								border-bottom: 1px solid #000000" align="left" valign=middle><font
									face="Times New Roman"><br></font></td>
							<td style="border-top: 1px solid #000000; border-left: 2px solid #000000;
								border-bottom: 1px solid #000000" align="left" valign=bottom><font
									face="Times New Roman"><br></font></td>
							<td style="border-top: 1px solid #000000; border-left: 2px solid #000000;
								border-bottom: 1px solid #000000" align="left" valign=bottom
								sdnum="1033;0;0"><font face="Times New Roman"><br></font></td>
							<td style="border-top: 1px solid #000000; border-left: 2px solid #000000;
								border-bottom: 1px solid #000000;border-right: 1px solid #000000" align="left" valign=bottom><font
									face="Times New Roman"><br></font></td>
							<td style="border-top: 1px solid #000000; border-left: 1px solid
								#000000;border-right: 2px solid #000000;
								border-bottom: 1px solid #000000" align="right" valign=bottom
								sdnum="1033;0;_(* #,##0_);_(* \(#,##0\);_(* &quot;-&quot;??_);_(@_)"><b><font
										face="Times New Roman"><br></font></b></td>
						</tr>
						<tr>
							<td style="border-top: 1px solid #000000; border-bottom: 1px solid #000000;
								border-left: 2px solid #000000; border-right: 1px  #000000"
								height="18" align="left" valign=bottom><b><font face="Times New Roman">Tax
										Amount in Words</font></b></td>
							<td style="border-top: 1px solid #000000; border-bottom: 1px solid #000000;
								border-left: 2px solid #000000; border-right: 2px solid #000000" colspan=7
								align="left" valign=bottom><font face="Times New Roman">{Igstvalue_words}</font></td>
						</tr>
						<tr>
							<td style="border-left: 2px solid #000000;border-top: 1px solid #000000"
								height="20" align="left"
								valign=bottom><b><font face="Times New Roman" color="#000000">Bank Account
										Details</font></b></td>
							<td style="border-top: 1px solid #000000" align="left" valign=bottom><font
									face="Times New Roman" color="#000000"><br></font></td>
							<td style="border-top: 1px solid #000000" align="left" valign=bottom><font
									face="Times New Roman" color="#000000"><br></font></td>
							<td style="border-top: 1px solid #000000" align="left" valign=bottom><font
									face="Times New Roman"><br></font></td>
							<td style="border-top: 1px solid #000000;border-right: 2px solid #000000"
								colspan=4 align="center"
								valign=bottom><b><font face="Times New Roman">for Object Technology
										Solutions India Pvt Ltd</font></b></td>
						</tr>
						<tr>
							<td style="border-left: 2px solid #000000" height="20" align="left"
								valign=bottom><font face="Times New Roman" color="#000000">Name of the Bank
									:</font></td>
							<td align="left" valign=bottom><font face="Times New Roman" color="#000000">HDFC
									Bank Ltd</font></td>
							<td align="left" valign=bottom><font face="Times New Roman" color="#000000"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="center" valign=bottom><b><font face="Times New Roman"><br></font></b></td>
							<td align="center" valign=bottom><b><font face="Times New Roman"><br></font></b></td>
							<td align="center" valign=bottom><b><font face="Times New Roman"><br></font></b></td>
							<td style="border-right: 2px solid #000000" align="center" valign=bottom><b><font
										face="Times New Roman"><br></font></b></td>
						</tr>
						<tr>
							<td style="border-left: 2px solid #000000" height="20" align="left"
								valign=bottom><font face="Times New Roman" color="#000000">Branch :</font></td>
							<td align="left" valign=bottom><font face="Times New Roman" color="#000000">Hitech
									City, Hyderabad</font></td>
							<td align="left" valign=bottom><font face="Times New Roman" color="#000000"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="center" valign=bottom><b><font face="Times New Roman"><br></font></b></td>
							<td align="center" valign=bottom><b><font face="Times New Roman"><br></font></b></td>
							<td align="center" valign=bottom><b><font face="Times New Roman"><br></font></b></td>
							<td style="border-right: 2px solid #000000" align="center" valign=bottom><b><font
										face="Times New Roman"><br></font></b></td>
						</tr>
						<tr>
							<td style="border-left: 2px solid #000000" height="20" align="left"
								valign=bottom><font face="Times New Roman" color="#000000">Account Number :</font></td>
							<td align="left" valign=bottom sdval="50200009926612" sdnum="1033;0;0"><font
									face="Times New Roman" color="#000000">50200009926612</font></td>
							<td align="left" valign=bottom sdnum="1033;0;0"><font face="Times New Roman"
									color="#000000"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td style="border-right: 2px solid #000000" align="left" valign=bottom><font
									face="Times New Roman"><br></font></td>
						</tr>
						<tr>
							<td style="border-left: 2px solid #000000" height="20" align="left"
								valign=bottom><font face="Times New Roman" color="#000000">IFSC :</font></td>
							<td align="left" valign=bottom><font face="Times New Roman" color="#000000">HDFC0000545</font></td>
							<td align="left" valign=bottom><font face="Times New Roman" color="#000000"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td style="width: 140px;" align="left" valign=bottom><b><font face="Times
										New Roman">Authorised signature</font></b></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td style="border-right: 2px solid #000000" align="left" valign=bottom><font
									face="Times New Roman"><br></font></td>
						</tr>
						<tr>
							<td style="border-left: 2px solid #000000" height="20" align="left"
								valign=bottom><font face="Times New Roman" color="#000000">SWIFT Code :</font></td>
							<td align="left" valign=bottom><font face="Times New Roman" color="#000000">HDFCINBBHYD</font></td>
							<td align="left" valign=bottom><font face="Times New Roman" color="#000000"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td style="border-right: 2px solid #000000" align="left" valign=bottom><font
									face="Times New Roman"><br></font></td>
						</tr>
						<tr>
							<td style="border-bottom: 1px solid #000000; border-left: 2px solid #000000"
								height="18" align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td style="border-bottom: 1px solid #000000" align="left" valign=bottom><font
									face="Times New Roman"><br></font></td>
							<td style="border-bottom: 1px solid #000000" align="left" valign=bottom><font
									face="Times New Roman"><br></font></td>
							<td style="border-bottom: 1px solid #000000" align="left" valign=bottom><font
									face="Times New Roman"><br></font></td>
							<td style="border-bottom: 1px solid #000000" align="left" valign=bottom><font
									face="Times New Roman"><br></font></td>
							<td style="border-bottom: 1px solid #000000" align="left" valign=bottom><font
									face="Times New Roman"><br></font></td>
							<td style="border-bottom: 1px solid #000000" align="left" valign=bottom><font
									face="Times New Roman"><br></font></td>
							<td style="border-bottom: 1px solid #000000; border-right: 2px solid
								#000000" align="left" valign=bottom><font face="Times New Roman"><br></font></td>
						</tr>
						<tr>
							<td style="border-top: 1px solid #000000; border-bottom: 1px solid #000000;
								border-left: 2px solid #000000; border-right: 2px solid #000000" colspan=8
								rowspan=2 height="35" align="left" valign=top><font face="Times New Roman">Decleration
									: We hereby declare that though our aggregate turnover in any preceding
									financial year from 2017-18 onwards is more than the aggregate turnover
									notified under sub-rule (4) of rule 48, we are not required to prepare an
									invoice in terms of the provisions of the said sub-rule</font></td>
						</tr>
						<tr>
							<td height="17" align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
						</tr>
						<tr>
							<td style="border-top: 1px solid #000000" height="27" width="27" colspan="2"
								align="left"
								valign=middle><b><font size=1 color="#06456B">USA . INDIA . COSTA RICA .
										DUBAI</font></b></td>
							<td style="border-top: 1px solid #000000" align="left" valign=middle><b><font
										size=1 color="#06456B"><br></font></b></td>
							<td style="border-top: 1px solid #000000" colspan=2 align="left"
								valign=bottom><font face="Tahoma" size=1 color="#231F20"># H-02, Phoenix
									Infocity SEZ, <br>Hitech City - 2 </br></font></td>
							<td style="border-top: 1px solid #000000" align="left" valign=bottom><font
									face="Times New Roman">Tel: +91 40 4425 1111 </font></td>
							<td style="border-top: 1px solid #000000" align="center" valign=bottom><font
									face="Times New Roman"><br></font></td>
							<td style="border-top: 1px solid #000000" align="center" valign=bottom><font
									face="Times New Roman">CIN:U72200TG2010PTC068549</font></td>
							<td style="border-top: 1px solid #000000" align="center" valign=bottom><font
									face="Times New Roman"><br></font></td>
						</tr>
						<tr>
							<td height="17" align="left" valign=bottom colspan="2"><font face="Times New
									Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td colspan="2" align="left" valign=bottom><font face="Tahoma" size=1
									color="#231F20">Hyderabad - 500081, India</font></td>
							<td align="left" valign=bottom><font face="Times New Roman">Fax: +91 40 4425
									1122</font></td>
							<td colspan=3 align="right" valign=middle><u><font color="#0000FF"><a
											href="mailto:info@otsi.co.in">E-mail: info@otsi.co.in</a></font></u></td>
						</tr>
						<tr>
							<td height="17" align="left" valign=middle><font face="Tahoma" size=1
									color="#231F20"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td colspan=2 align="right" valign=bottom><u><font color="#0000FF"><a
											href="https://otsi-global.com/">https://otsi-global.com/</a></font></u></td>
						</tr>
						<tr>
							<td height="17" align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
						</tr>
						<tr>
							<td height="17" align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
							<td align="left" valign=bottom><font face="Times New Roman"><br></font></td>
						</tr>
					</table>
			</div>
				</body>
			</html>
		"""
listofhtmls = []
output_pdf_filename = []
pdfkit_html_pdf_config_filepath = 'C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe'
uploads_file_path = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "upload")
pdf_file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pdf")
lut_folder_path = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "LUT")


def money_in_words(amount_num):
    """Funtion to get the amount in words (Indian Currancy Format)."""
    log.info(
        f"Calling 'money_in_words' function to get the amount in words. amount given - {amount_num}")
    amount = round(float(amount_num), 2)
    money = ""
    decimals = ""
    if int(amount) != 0:
        if int(amount) == 1:
            money = num2words(int(amount), lang='en_IN').replace(
                ",", "")+" rupee "
        else:
            money = num2words(int(amount), lang='en_IN').replace(
                ",", "")+" rupees "
    if int(str(amount).split(".")[-1]) != 0:
        decimals = num2words(
            int(str(amount).split(".")[-1]), lang='en_IN')+" paisa"
    if (money == "") & (decimals == ""):
        log.info(f"Amount in words - Nil")
        return "Nil"
    amount_in_string = (money+decimals+" only").replace("-",
                                                        " ").replace("  ", " ").title().replace("One Rupees", "One Rupee")
    log.info(f"Amount in words - {amount_in_string}")
    return amount_in_string


filename = str(args.filename).strip()
log.info(f"Validating the source file format.")
if filename.split(".")[-1] != "xlsx":
    log.critical(
        f"The source file which you are referring {filename} is not Excel, Only excel with .xlsx extension is allowed.")
    raise InvalidFileFormat('Only excel with .xlsx extension is allowed.')

try:
    log.info(f"Creating Pandas dataframe from {filename}")
    # Creating a Pandas dataframe.
    df = pd.read_excel(uploads_file_path+"//"+filename)
    # Removing trail whitespaces in the column names.
    df.columns = [col.strip() for col in df.columns]
    # Rounding of the numerical columns to 2 decimals.
    log.info("Rounding of the numerical columns to 2 decimals.")
    df[['Bill Rate', 'Sum of Invoice Line Item Amount', 'Sum of Less MSP Fee & VMS Fee 2.6%', 'Sum of Total Invoice Amnt without GST', 'Sum of Total Invoice Amnt']] = df[['Bill Rate',
                                                                                                                                                                           'Sum of Invoice Line Item Amount', 'Sum of Less MSP Fee & VMS Fee 2.6%', 'Sum of Total Invoice Amnt without GST', 'Sum of Total Invoice Amnt']].applymap(lambda val: round(float(val), 2))
    # Standardizing the date format.
    log.info("Standardizing the date format.")
    df['Invoice Date'] = df['Invoice Date'].dt.strftime('%d-%m-%Y')
    df['WO Date'] = df['WO Date'].dt.strftime('%d-%m-%Y')
    log.info("Calculating GST value based on total invoice amount and GST percentage.")
    # Calculating GST value based on total invoice amount and GST percentage.
    df['gst_value'] = df['Sum of Total Invoice Amnt without GST'] * df['GST']
    # Calling function to convert the amount in words (Indian Currancy Format).
    log.info(
        "Calling function to convert the amount in words (Indian Currancy Format).")
    df['sumttotal_words'] = df['Sum of Total Invoice Amnt'].apply(
        money_in_words)
    df['Igstvalue_words'] = df['gst_value'].apply(money_in_words)
    # Bringing all the columns to string datatype.
    log.info("Bringing all the columns to string datatype.")
    df[df.columns] = df[df.columns].astype(str)
    # Getting each row in the pandas dataframe in Tuples.
    log.info("Getting each row in the pandas dataframe in Tuples.")
    invoice_info = [tuple(row) for row in df.itertuples(index=False)]
except Exception as e:
    log.critical(
        f"An error occure while reading and analysis the source data. Error: {e}")

try:
    print("Generating Dynamic HTML.")
    log.info("Iterating loop to generate dynamic HTML.")
    # Loop to generate a HTML content using invoice information.
    for inv_no, inv_prd, inv_dt, wonum, wodt, rvcrg, ciid, pos, sta, gstin, bta, bgstin, wrkr_name, itmdisc, bhr, brt, silia, slmspfee, stiawgst, gst, sumta, gst_value, sumttotal_words, Igstvalue_words in tqdm(invoice_info):
        listofhtmls.append(html_header + html_styling + html_body.format(inv_no=inv_no, inv_prd=inv_prd, inv_dt=inv_dt, wonum=wonum, wodt=wodt, rvcrg=rvcrg, ciid=ciid, pos=pos, sta=sta, gstin=gstin, bta=bta, bgstin=bgstin,
                           wrkr_name=wrkr_name, itmdisc=itmdisc, bhr=bhr, brt=brt, silia=silia, slmspfee=slmspfee, stiawgst=stiawgst, gst=gst, sumta=sumta, gst_value=gst_value, sumttotal_words=sumttotal_words, Igstvalue_words=Igstvalue_words))
        output_pdf_filename.append(inv_no)
except Exception as e:
    log.critical(
        f"An error occure while iterating loop to generate dynamic HTML. Error: {e}")

# Configuring options to manage PDF page dimension.
options = {
    'page-size': 'A4',
    'margin-top': '0.5in',
    'margin-right': '0.5in',
    'margin-bottom': '0.5in',
    'margin-left': '0.5in'
}

try:
    log.info(
        "Initialising and configuring the pdfkit library to create PDF using HTML content.")
    # Initialising and configuring the pdfkit library to create PDF using HTML content.
    config = pdfkit.configuration(wkhtmltopdf=pdfkit_html_pdf_config_filepath)
except Exception as e:
    log.critical(
        f"An error occure while configuring pdfkit library to create PDF using html. Error: {e}")

# Handling exceptions which creating invoice PDF's using HTML content.
try:
    log.info("Creating PDFs using HTML.")
    # Reading filesystem paths with semantics appropriate for different operating systems.
    # Create folder if not exists in the system path.
    pdf_folder_dir = pdf_file_path+"/"+filename.split(".")[0].replace(
        " ", "_")+"_"+datetime.now().strftime("%d_%m_%Y_%H_%M_%S")+"/"
    output_dir = Path(pdf_folder_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    print("Creating PDF using HTML Generated Above.")
    # Loop to iterate invoice's HTML contents.
    for html_page, pdf_file_name in tqdm(list(zip(listofhtmls, output_pdf_filename))):
        # Converting and saving pdf using pdfkit.
        pdfkit.from_string(html_page, pdf_folder_dir+pdf_file_name +
                           ".pdf", options=options, configuration=config)
except Exception as e:
    log.critical(
        f"An error occure while creating PDFs using HTML string. Error: {e}")

# Validating the LUT file extension.
log.info("Validating the LUT file extension.")
lut_filename = str(args.lut).strip()
if lut_filename.split(".")[-1] != "pdf":
    log.critical(
        f"The file which you have uploaded {lut_filename} is not valid, LUT file with .pdf extension is only allowed.")
    raise InvalidFileFormat('Only pdf with .pdf extension is allowed.')

# Handling exceptions while merge invoice PDF with LUT PDF.
try:
    log.info("Appending LUT file at the bottom of each Invoice.")
    # Open LUT PDF file in read mode.
    lut_file = open(lut_folder_path+'/'+lut_filename, 'rb')
    lut_file_reader = PyPDF2.PdfReader(lut_file)
    print("Appending LUT forms to the invoice PDFs generated above.")
    # List of pdf files in a folder.
    for invoice_pdf in tqdm(os.listdir(pdf_folder_dir)):
        # Initialising PDF merger object.
        pdfMerge = PyPDF2.PdfMerger()
        # Open PDF file in read mode only.
        invoice_File = open(pdf_folder_dir+invoice_pdf, 'rb')
        # Read the PDF using pdfReader method.
        invoice_reader = PyPDF2.PdfReader(invoice_File)
        # Merge the PDF's in list.
        pdfMerge.append(invoice_reader)
        pdfMerge.append(lut_file_reader)
        # Close the PDF file.
        invoice_File.close()
        # Write the PDF to defined path.
        pdfMerge.write(pdf_folder_dir+invoice_pdf)
    lut_file.close()
except Exception as e:
    log.critical(
        f"An error occure while appending LUT file at the bottom of invoice PDFs. Error: {e}")
