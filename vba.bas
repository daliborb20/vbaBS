
Option Explicit

Sub bilansStanja()

    
Application.ScreenUpdating = False



    
'##############zaglavlje
Range("b1") = "Group of accounts, account"
Range("c1") = "ITEM"
Range("D1") = "ADP"
Range("E1") = "Note number"
Range("F1") = "Current year"
Range("g1") = "Previous year (closing balance)"
Range("h1") = "Previous year (opening balance)"
Range("c2") = "Assets"

'#############kolona b

Range("b7") = "011 and 012 and 014"
Range("B9") = "015 and 016"
Range("b15") = "025 and 027"
Range("b16") = "026 and 028"
Range("b20") = "04 and 05"
Range("b21") = "040 (part), 041(part) and 042"
Range("b22") = "040 (part), 041 (part) and 042"
Range("b23") = "043, 050 (part) and 051 (part)"
Range("b24") = "044, 050 (part) and 051 (part)"
Range("b25") = "045 (part) and 053 (part) "
Range("b26") = "045 (part) and 053 (part) "
Range("b30") = "028 (part), except 288"
Range("b33") = "1, except 14"
Range("b35") = "11 and 12"
Range("b37") = "150, 152 and 154"
Range("b38") = "151, 153 and 155"
Range("b43") = "200 and 202"
Range("b44") = "201 and 203"
Range("b46") = "21, 22 and 27"
Range("b47") = "21, 22 except 223 and 224, and 27"
Range("b53") = "232 and 234 (part)"
Range("b54") = "233 and 234 (part)"
Range("b56") = "236 (part)"
Range("b58") = " 236 (part), 238 and 239"
Range("b60") = "28 (part), except 288"
Range("b65") = "30, except 306"
Range("b60") = "330 and credit balance of 331,332,333,334,335,336,337"
Range("b70") = "debit balance of 331, 332, 333, 334, 335, 336, 337"
Range("b82") = "40 except 400 and 404"
Range("b85") = "411 (part) and 412 (part)"
Range("b86") = "411 (part) and 412 (part)"
Range("b87") = "414 and 416 (part)"
Range("b88") = "414 and 416 (part)"
Range("b91") = "49 (part), except 498 and 495 (part)"
Range("b93") = "495 (part)"
Range("b96") = "42, except 427"
Range("b97") = "420 (part) and 421 (part)"
Range("b98") = "420 (part) and 421 (part)"
Range("b99") = "422 (part), 424 (part), 425 (part) and 429 (part)"
Range("b100") = "422 (part), 424 (part), 425 (part) and 429 (part)"
Range("b101") = "423,  424 (part), 425 (part) and 429 (part)"
Range("b105") = "42 except 430"
Range("b106") = "431 and 433"
Range("b107") = "432 and 434"
Range("b110") = "439 (part)"
Range("b111") = "439 (part)"
Range("b112") = "44, 45, 46 except 467, 47 and 48"
Range("b113") = "44, 45, 46 except 467"
Range("b114") = "47, 48, except 481"
Range("b117") = "49 except 498"

Range("c2") = "Assets"
Range("C3") = "A. Subscribed capital unpaid"
Range("c4") = "B. Permanent assets (0003+0009+0017+0018+0028)"
Range("c5") = "I. Intangible assets (0004+0005+0006+0007+0008)"
Range("c6") = "1. Investment in development"
Range("c7") = "2. Concessions, patents, licencses, trademarks, service marks, software and similar rigths"
Range("c8") = "3. Goodwill"
Range("c9") = "4. Intangible assets in finance lease and intangible assets in preparation"
Range("c10") = "5. Advances for intangible assets"
Range("c11") = "II. Immovables, plants and equipment (0011+0012+0013+0014+0015+0016+0017+0018)"
Range("c12") = "1. Land and buildings"
Range("c13") = "2. Plant and equipment"
Range("c14") = "3. Investment immovables"
Range("c15") = "4. PPE in financial lease and PPE in preparation"
Range("c16") = "5. Other PPE and investments in third-party PPE"
Range("c17") = "6. Advances for PPE - domestic"
Range("c18") = "7. Advances for PPE - foreign"
Range("c19") = "III. Bilogical resources"
Range("c20") = "IV. Long-term financial investments and long-term financial receivables (0019+0020+0021+0022+0023+0024+0025+0026+0027)"
Range("c21") = "1. Investments in capital of legal entities (non-equity method)"
Range("c22") = "2. Investments in capital of legal entities (equity method)"
Range("c23") = "3. Long-term investments in parent, subsidiaries and other associated companies - domestic"
Range("c24") = "4. Long-term investments in parent, subsidiaries and other associated companies - foreign"
Range("c25") = "5. Long-term domestic loans"
Range("c26") = "6. Long-term foreign loans"
Range("c27") = "7. Securities (amortised cost calculated)"
Range("c28") = "8.Own shares purchased"
Range("c29") = "9. Other long-term investments and other long-term receivables"
Range("c30") = "V. Accured expenses (long term)"
Range("c31") = "VI. Deferred tax assets"

'############obrtna imovina

Range("c32") = "G. Current assets (0031+0037+0038+0044+0048+0057+0058)"
Range("c33") = "I. Inventories (0032+0033+0034+0035+0036)"
Range("c34") = "1. Materials, spare parts, tools and small inventory"
Range("c35") = "2. Work and services in progress and finished products"
Range("c36") = "3. Goods"
Range("c37") = "4. Advances paid for inventories and services - domestic "
Range("c38") = "5. Advances paid for inventories and services - foreign "
Range("c39") = "II. Permanent assets hold for sale"
Range("c40") = "III. Receivables from sales (0052+0053+0054+0055+0056+0057+0058)"
Range("c41") = "1. Trade receivables - domestic"
Range("c42") = "2. Trade receivables - foreign"
Range("c43") = "3. Trade receivables from parent, subsidiares and other associated companies - domestic"
Range("c44") = "4. Trade receivables from parent, subsidiares and other associated companies - foreign"
Range("c45") = "5. Other trade receivables"
Range("c46") = "IV. Other short-term receivables"
Range("c47") = "1. Other receivables (0045+0046+0047)"
Range("c48") = "2. Prepaid company income tax"
Range("c49") = "3. Receivables from prepaid other taxes and social contributions"
Range("c50") = "V. Short-term financial investments (0049+0050+0051+0052+0053+0054+0055+0056)"
Range("c51") = "1. Short-term loans and investments in parent companies and subsidiaries"
Range("c52") = "2. Short-term loans and investments in other associated companies"
Range("c53") = "3. Short-term loans and investments - domestic"
Range("c54") = "4. Short-term loans and investments - foreign"
Range("c55") = "5. Securities (amortised cost calculated)"
Range("c56") = "6. Financial assets at fair value through profit or loss"
Range("c57") = "7. Own shares purchased"
Range("c58") = "8. Other short-term financial investments"
Range("c59") = "VI. Cash and cash equivalents"
Range("c60") = "VII. Accured expenses (short-term)"
Range("c61") = "D. Total assets = Operating assets (0001+0002+0029+0030)"
Range("c62") = "Dj. Off-balance assets"
Range("c63") = "EQUITY AND LIABILITIES"
'#########################kapital

Range("c64") = "A. Equity (0402+0403+0404+0405+0406-0407+0408+0411-0412) >= 0"
Range("c65") = "I. Share capital"
Range("c66") = "II. Subscribed capital unpaid"
Range("c67") = "III. Share premium"
Range("c68") = "IV. Reserves"
Range("c69") = "V. Positive revaluation reserves and equipment unrealized profits from securities and other elements of comprehensive income"
Range("c70") = "VI. Unrealized losses from securities and other comprehensive income"
Range("c71") = "VII. Retained earnings (0409+0410)"
Range("c72") = "1. Retained earnings for the current year"
Range("c73") = "2. Retained earnings for the current year"
Range("c74") = "Participation without control rights"
Range("c75") = "IX. Loss"
Range("c76") = "1. Loss from the previous years"
Range("c77") = "2. Loss from the current year"
Range("c78") = "B. Long-term provisions and liabilities (0416+0420+0428)"
Range("c79") = "I. Long-term provisions (0417+0418+0419)"
Range("c80") = "1. Provisions from compensations and other employment benefits"
Range("c81") = "2. Provisions for costs incurred during warranty period"
Range("c82") = "3. Other long-term provisions"
Range("c83") = "II. Long term liabilities (0421+0422+0423+0424+0425+0426+0427)"
Range("c84") = "1. Debts convertible into equity"
Range("c85") = "2. Long-term loans and other liabilities to parent, subsidiaries and other associated companies - domestic"
Range("c86") = "3. Long-term loans and other liabilities to  liabilities to parent, subsidiaries and other associated companies - foreign"
Range("c87") = "4. Long-term credits, loans and other liabilities - domestic"
Range("c88") = "5. Long-term credits, loans and other liabilities - foreign"
Range("c89") = "6. liabilities for long-term securities "
Range("c90") = "7. Other long-term liabilities"
Range("c91") = "III. Deferred expenses (long-term)"
Range("c92") = "V. Deferred tax liabilities"
Range("c93") = "G. Long-term deferred income and received grants"
Range("c94") = " D. Short-term provisions and liabilities (0432+0433+0441+0442+0449+0453+0454)"
Range("c95") = "I. Short-term provisions"
Range("c96") = "II. Short-term liabilities (0434+0435+0436+0437+0438+0439+0440)"
Range("c97") = "Loans from parent, subidiaries and associated companies - domestic"
Range("c98") = "Loans from parent, subidiaries and associated companies - foreign"
Range("c99") = "Loans from legal entities that are not banks"
Range("c100") = "Loans from domestic banks"
Range("c101") = "Loans and other obligations - foreign"
Range("c102") = "Liabilities for short-term securities"
Range("c103") = "Liabilities for financial derivatives"
Range("c104") = "III. Prepayments, deposits and guarantees"
Range("c105") = "IV. Operating liabilities (0443+0444+0445+0446+0447+0448)"
Range("c106") = "1. Trade payables - domestic parent, subsidiaries, and associated companies"
Range("c107") = "2. Trade payables - foreign parent, subsidiaries, and associated companies"
Range("c108") = "3. Trade payables - domestic"
Range("c109") = "4. Trade payables - foreign"
Range("c110") = "5. Liabilities for bill of exchange"
Range("c111") = "6. Other liabilities"
Range("c112") = "V. Other short-term liabilities (0450+0451+0452)"
Range("c113") = "1. Other short-term liabilities"
Range("c114") = "2. Liabilities for the VAT and other taxes"
Range("c115") = "3. Liabilities for company income tax"
Range("c116") = "VI. Liabilities for permanent assets held for sale"
Range("c117") = "VII Deferred expenses (short-term)"
Range("c118") = "Loss above equity (0415+0429+0430+0431-0059) >=0 = (0407+0412-0402-0403-0404-0405-0406-0408-0411) >= 0"
Range("c119") = "Total equity and liabilities (0401+0415+0429+0430+0431-0455) >= 0"
Range("c120") = "Off-balance sheet liabilities"


'####################setovanje kolona i redova
Cells.Select
With Selection
    .WrapText = True
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .MergeCells = False
End With
Cells.EntireRow.AutoFit
Columns("F:H").Select
Selection.ColumnWidth = 17
Columns("c").Select
Selection.ColumnWidth = 55
Rows("120:1").EntireRow.AutoFit
Columns("b").ColumnWidth = 20
ActiveWindow.DisplayGridlines = False
Columns("h").Delete

ActiveSheet.PageSetup.PrintArea = Range("a1:g120")


Application.ScreenUpdating = True

Range("a1").Activate

End Sub




Sub bilansUspeha()

Application.ScreenUpdating = False


'##############zaglavlje
Range("b1") = "Group of accounts, account"
Range("c1") = "ITEM"
Range("D1") = "ADP"
Range("E1") = "Note number"
Range("F1") = "Current year"
Range("g1") = "Previous year (closing balance)"

'#########colona b

Range("b5") = "600, 602 and 604"
Range("b6") = "601, 603 and 605"
Range("b8") = "610, 612 and 614"
Range("b9") = "611, 613 and 615"
Range("b13") = "64 and 65"
Range("b14") = "68, except 683, 685 and 686"
Range("b22") = "52 except 520 and 521"
Range("b24") = "58 except 583, 585, 586"
Range("b26") = "54 except 540"
Range("b31") = "660 and 661"
Range("b33") = "663 and 664"
Range("b34") = "665 and 669"
Range("b36") = "560 and 561"
Range("b39") = "565 and 569"
Range("b42") = "683, 685 and 686"
Range("b43") = "583, 585 and 586"
Range("b56") = "722 debit balance"
Range("b57") = "722 credit balance"

'#########Colona C
Range("c2") = "INCOME FROM OPERATIONS"
Range("c3") = "A. OPERATING INCOME (1002+1005+1008+1009-1010+1011+1012)"
Range("c4") = "I. INCOME FROM GOODS SOLD (1003+1004)"
Range("c5") = "1. Income from goods sold - domestic"
Range("c6") = "2. Income from goods sold - foreign"
Range("c7") = "II. INCOME FROM PRODUCTS SOLD AND SERVICES PROVIDED"
Range("c8") = "1. Products sold and services provided to domestic customers"
Range("c9") = "2. Products sold and services provided to foreign customers"
Range("c10") = "III. REVENUE FROM UNDERTAKING FOR OWN PURPOSES"
Range("c11") = "IV. INCREASE IN INVENTORIES OF WORK IN PROGRESS AND FINISHED PRODUCTS AND UNFINISHED SERVICES"
Range("c12") = "V. DECREASE IN INVENTORIES OF WORK IN PROGRESS AND FINISHED PRODUCTS AND UNFINISHED SERVICES"
Range("c13") = "VI. OTHER OPERATING INCOME "
Range("c14") = "VII. INCOME FROM VALUE ADJUSTING OF ASSETS (EXCEPT FINANCIAL ASSETS)"
Range("c15") = "EXPENSES FROM OPERATIONS"
Range("c16") = "B. OPERATING EXPENSES (1014+1015+1016+1020+1021+1022+1023+1024)"
Range("c17") = "I. COST OF GOODS SOLD"
Range("c18") = "II. RAW MATERIAL, FUEL AND ENERGY COSTS"
Range("c19") = "III. SALARIES, WAGES, AND OTHER PERSONAL INDEMNITIES"
Range("c20") = "1. Cost of salaries and fringe benefits"
Range("c21") = "2.Costs of taxes and contributions on salaries and fringe benefits"
Range("c22") = "3. Other personal indemnities "
Range("c23") = "IV. DEPRECIATION COSTS"
Range("C24") = "V. COSTS FROM VALUE ADJUSTMENT OF ASSETS (EXCEPT FINANCIAL ASSETS)"
Range("c25") = "VI. PRODUCTION SERVICES COSTS"
Range("c26") = "VII. PROVISION COSTS"
Range("c27") = "VIII. INTANGIBLE COSTS"
Range("c28") = "C. OPERATING PROFIT (1001-1013) >= 0"
Range("c29") = "D. OPERATING LOSS (1013-1001) >= 0"
Range("c30") = "E. FINANCIAL INCOME (1028+1029+1030+1031)"
Range("c31") = "F. FINANCIAL INCOME FROM PARENT, SUBSIDIARIES AND ASSOCIATED COMPANIES"
Range("c32") = "II. INCOME FROM INTEREST"
Range("c33") = "III. POSITIVE EFFECTS ON EXCHANGE RATE AND EFFECTS OF FOREIGN CURRENCY CLAUSE"
Range("c34") = "IV. OTHER FINANCIAL INCOME"
Range("c35") = "G. FINANCIAL EXPENSES (1033+1034+1035+1036)"
Range("c36") = "H. FINANCIAL EXPENSES INCURRED WITH PARENT, SUBSIDIARIES AND ASSOCIATED COMPANIES"
Range("c37") = "II. INTEREST EXPENSES"
Range("c38") = "III. NEGATIVE EFFECTS ON EXCHANGE RATE AND EFFECTS OF FOREIGN CURRENCY CLAUSE "
Range("c39") = "IV. OTHER FINANCIAL EXPENSES"
Range("c40") = "I. PROFIT FROM FINANCING (1027-1032)"
Range("c41") = "K. LOSS FROM FINANCING (1032-1027)"
Range("c42") = "L. INCOME ON VALUE ADJUSTMENT OF OTHER ASSETS CARRIED AT FAIR VALUE THROUGH PROFIT AND LOSS ACCOUNT"
Range("c43") = "M. EXPENSES ON VALUE ADJUSTMENT OF OTHER ASSETS CARRIED AT FAIR VALUE THROUGH PROFIT AND LOSS ACCOUNT "
Range("c44") = "N. OTHER INCOME"
Range("c45") = "O. OTHER EXPENSES"
Range("c46") = "P. TOTAL INCOME (1001+10027+1039+1041)"
Range("c47") = "Q. TOTAL EXPENSES (1013+1032+1040+1042) "
Range("c48") = "R. PROFIT FROM REGULAR OPERATIONS BEFORE TAX (1043-1044) >= 0"
Range("c49") = "S. LOSS FROM REGULAR OPERATIONS BEFORE TAX (1044-1043) >= 0"
Range("c50") = "T. NET PROFIT FROM DISCONTINUED OPERATIONS, EFFECTS OF CHANGES IN ACCOUNTING POLICIES AND CORRECTIONS OF ERRORS FROM PREVIOUS PERIODS"
Range("c51") = "U.  NET LOSS FROM DISCONTINUED OPERATIONS, EFFECTS OF CHANGES IN ACCOUNTING POLICIES AND CORRECTIONS OF ERRORS FROM PREVIOUS PERIODS"
Range("c52") = "V. PROFIT BEFORE TAX (1045-1046+1047-1048)>=0"
Range("C53") = "W. LOSS BEFORE TAX (1046-1045+1048-1047) >=0 "
Range("c54") = "X. TAX ON PROFIT"
Range("c55") = "I. TAX EXPENSES FOR THE PERIOD"
Range("c56") = "II. DEFERRED TAX EXPENSES OF A PERIOD"
Range("c57") = "III. DEFERRED TAX INCOME OF A PERIOD"
Range("c58") = "Y. PERSONAL INDEMNITIES PAID TO EMPLOYER"
Range("c59") = "Z. NET PROFIT (1049-1050-1051-1052+1053-1054) >= 0"
Range("c60") = "Z1 NET LOSS (1050-1049+1051+1052-1053+1054) >= 0"
Range("c61") = "I. NET PROFIT WHICH BELONGS TO MINORITY INVESTORS"
Range("c62") = "II. NET PROFIT WHICH BELONGS TO MAJORITY OWNER"
Range("c63") = "III. NET LOSS WHICH BELONGS TO MINORITY INVESTORS"
Range("c64") = "IV. NET LOSS WHICH BELONGS TO MAJORITY OWNER"
Range("c65") = "V. EARNINGS PER SHARE"
Range("c66") = "1. Basic earnings per share"
Range("c67") = "2. Diluted earnings per share"


'####################setovanje kolona i redova
Cells.Select
With Selection
    .WrapText = True
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .MergeCells = False
End With
Cells.EntireRow.AutoFit
Columns("F:H").Select
Selection.ColumnWidth = 17
Rows("1:1").EntireRow.AutoFit
ActiveWindow.DisplayGridlines = False


Range("a1").Activate

Application.ScreenUpdating = True
End Sub





