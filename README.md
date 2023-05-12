
    function getLoadSodRiskGrid(data) {
        debugger;
        var workbook = new ExcelJS.Workbook();
        var worksheet = workbook.addWorksheet("Credit Card");
        var worksheet2 = workbook.addWorksheet("Normal");

        worksheet.mergeCells('C1', 'E1');
        worksheet.getCell('C1').value = 'E-Ledger Gen. Journal Line';
      
        worksheet.getCell('F1').value = '81';


        worksheet2.mergeCells('C1', 'E1');
        worksheet2.getCell('C1').value = 'E-Ledger Gen. Journal Line';

        worksheet2.getCell('F1').value = '81';


        const row = worksheet.getRow(5);
        '@DBResources.GetText("Journal Template Name")',
        row.values = ['@DBResources.GetText("Journal Template Name")', '@DBResources.GetText("Journal Batch Name")', '@DBResources.GetText("Line No.")', '@DBResources.GetText("Account Type")', '@DBResources.GetText("Account No.")', '@DBResources.GetText("Document Date")', '@DBResources.GetText("Posting Date")', '@DBResources.GetText("Document Type")', '@DBResources.GetText("Document No.")', '@DBResources.GetText("Description")', '@DBResources.GetText("Currency Code")', '@DBResources.GetText("Amount")', '@DBResources.GetText("Amount (LCY)")', '@DBResources.GetText("Currency Factor")', '@DBResources.GetText("External Document No")', '@DBResources.GetText("Invoice Source Type")'];

        const font = {
            name: 'Calibri',
            size: 12,
            bold: true
        };

        row.eachCell((cell) => {
            cell.font = font;
        });
        worksheet.columns = [
            { key: 'journalTemplateName', width: 20, style: { font: { name: 'Arial' } } },
            { key: 'journalBatchName', width: 20, style: { font: { name: 'Arial' } } },
            { key: 'lineNo', width: 20, style: { font: { name: 'Arial' } } },
            { key: 'accountType', width: 20, style: { font: { name: 'Arial' } } },
            { key: 'accountNo', width: 20, style: { font: { name: 'Arial' } } },
            { key: 'documentDate', width: 20, style: { font: { name: 'Arial' } } },
            { key: 'postingDate', width: 20, style: { font: { name: 'Arial' } } },
            { key: 'documentType', width: 20, style: { font: { name: 'Arial' } } },
            { key: 'documentNo', width: 20, style: { font: { name: 'Arial' } } },
            { key: 'description', width: 20, style: { font: { name: 'Arial' } } },
            { key: 'currencyCode', width: 20, style: { font: { name: 'Arial' } } },
            { key: 'amount', width: 20, style: { font: { name: 'Arial' } } },
            { key: 'amountLCY', width: 20, style: { font: { name: 'Arial' } } },
            { key: 'currencyFactor', width: 20, style: { font: { name: 'Arial' } } },
            { key: 'externalDocumentNo', width: 20, style: { font: { name: 'Arial' } } },
            { key: 'invoiceSourceType', width: 20, style: { font: { name: 'Arial' } } },

        ];

        const row2 = worksheet2.getRow(5);

        row2.values = ['@DBResources.GetText("Journal Template Name")', '@DBResources.GetText("Journal Batch Name")', '@DBResources.GetText("Line No.")', '@DBResources.GetText("Account Type")', '@DBResources.GetText("Account No.")', '@DBResources.GetText("Document Date")', '@DBResources.GetText("Posting Date")', '@DBResources.GetText("Document Type")', '@DBResources.GetText("Document No.")', '@DBResources.GetText("Description")', '@DBResources.GetText("Currency Code")', '@DBResources.GetText("Amount")', '@DBResources.GetText("Amount (LCY)")', '@DBResources.GetText("Currency Factor")', '@DBResources.GetText("External Document No")', '@DBResources.GetText("Invoice Source Type")'];

        row2.eachCell((cell) => {
            cell.font = font;
        });
        worksheet2.getRow(5).columns = [
            { key: 'journalTemplateName', width: 20, style: { font: { name: 'Arial' } } },
            { key: 'journalBatchName', width: 20, style: { font: { name: 'Arial' } } },
            { key: 'lineNo', width: 20, style: { font: { name: 'Arial' } } },
            { key: 'accountType', width: 20, style: { font: { name: 'Arial' } } },
            { key: 'accountNo', width: 20, style: { font: { name: 'Arial' } } },
            { key: 'documentDate', width: 20, style: { font: { name: 'Arial' } } },
            { key: 'postingDate', width: 20, style: { font: { name: 'Arial' } } },
            { key: 'documentType', width: 20, style: { font: { name: 'Arial' } } },
            { key: 'documentNo', width: 20, style: { font: { name: 'Arial' } } },
            { key: 'description', width: 20, style: { font: { name: 'Arial' } } },
            { key: 'currencyCode', width: 20, style: { font: { name: 'Arial' } } },
            { key: 'amount', width: 20, style: { font: { name: 'Arial' } } },
            { key: 'amountLCY', width: 20, style: { font: { name: 'Arial' } } },
            { key: 'currencyFactor', width: 20, style: { font: { name: 'Arial' } } },
            { key: 'externalDocumentNo', width: 20, style: { font: { name: 'Arial' } } },
            { key: 'invoiceSourceType', width: 20, style: { font: { name: 'Arial' } } },

        ];
        var lineNo = 0;

        for (var i = 0; i < data.data.length; i++) {


                if (data.data[i]['PaymentType'] == "PaymentType02") {

                    var docNo = data.data[i]["DocumentNo"];
                    var accountNo = data.data[i]["ExpenseNameId"];
                    var docDate = data.data[i]["DocumentDate"];
                    var postDate = data.data[i]["Date"];
                    var description = data.data[i]["Description"];

                    lineNo += 10000;
                    var journaltemplate;
                    var accountType;
                    var docType;
                    var currencyCode;
                    var expenseAmount = data.data[i]["ExpenseAmount"];
                    var fxRate = data.data[i]["FxRate"];
                    var amountLCY = expenseAmount * fxRate;
                    var billNo = data.data[i]["BillNo"].length;
                    var invoiceSourceType;

                    if (billNo == 16) {
                        invoiceSourceType = "Electronic";
                    }
                    else {
                        invoiceSourceType = "Paper";
                    }

                    if (data.data[i]["CurrencyType"] == "TRY") {
                        currencyCode = " ";
                    }
                    else {
                        currencyCode = data.data[i]["CurrencyType"];
                    }

                    if (data.data[i]['DocumentType'] == "DocumentType02") {
                        journaltemplate = '1-INVOICE';
                        docType = 'INVOICE';

                    }
                    else {
                        journaltemplate = '1-MASRAF';
                        docType = 'Payment';

                    }
                    if (data.data[i]["CurrencyType"] != "TRY") {
                        currencyFactor = 1 / data.data[i]["FxRate"];

                    }
                    else {
                        currencyFactor = 0;
                    }



                    worksheet.addRow({ journalTemplateName: journaltemplate, journalBatchName: docNo, lineNo: lineNo, accountType: 'GL/Account', accountNo: '770.02.002', documentDate: docDate, postingDate: postDate, documentType: docType, documentNo: 'YDMH-' + docNo, description: description, currencyCode: currencyCode, amount: expenseAmount, amountLCY: amountLCY, currencyFactor: currencyFactor, externalDocumentNo: docNo, invoiceSourceType: invoiceSourceType });

                }
                else if (data.data[i]['PaymentType'] == "PaymentType01") {
                    //worksheet2.addRow({ UserName: data.data[i].PaymentType, transactionCodeI: data.data[i].PaymentType });
                    var docNo = data.data[i]["DocumentNo"];
                    var accountNo = data.data[i]["ExpenseNameId"];
                    var docDate = data.data[i]["DocumentDate"];
                    var postDate = data.data[i]["Date"];
                    var description = data.data[i]["Description"];

                    lineNo += 10000;
                    var journaltemplate;
                    var accountType;
                    var docType;
                    var currencyCode;
                    var expenseAmount = data.data[i]["ExpenseAmount"];
                    var fxRate = data.data[i]["FxRate"];
                    var amountLCY = expenseAmount * fxRate;
                    var billNo = data.data[i]["BillNo"].length;
                    var invoiceSourceType;

                    if (billNo==16) {
                        invoiceSourceType = "Electronic";
                    }
                    else {
                        invoiceSourceType = "Paper";
                    }

                    if (data.data[i]["CurrencyType"] == "TRY") {
                        currencyCode = " ";
                    }
                    else {
                        currencyCode = data.data[i]["CurrencyType"];
                    }

                    if (data.data[i]['DocumentType'] == "DocumentType02") {
                        journaltemplate = '1-INVOICE';
                        docType = 'INVOICE';

                    }
                    else {
                        journaltemplate = '1-MASRAF';
                        docType = 'Payment';

                    }
                    if (data.data[i]["CurrencyType"] != "TRY") {
                        currencyFactor = 1 / data.data[i]["FxRate"];

                    }
                    else {
                        currencyFactor = 0;
                    }



                    worksheet2.addRow({ journalTemplateName: journaltemplate, journalBatchName: docNo, lineNo: lineNo, accountType: 'GL/Account', accountNo: '770.02.002', documentDate: docDate, postingDate: postDate, documentType: docType, documentNo: 'YDMH-' + docNo, description: description, currencyCode: currencyCode, amount: expenseAmount, amountLCY: amountLCY, currencyFactor: currencyFactor, externalDocumentNo: docNo, invoiceSourceType: invoiceSourceType });

                    debugger;

                }



        }
        for (var i = 0; i < data.data.length; i++) {


            if (data.data[i]['PaymentType'] == "PaymentType02") {
                var docNo = data.data[i]["DocumentNo"];
                var docDate = data.data[i]["DocumentDate"];
                var postDate = data.data[i]["Date"];
                var description = data.data[i]["Description"];
                var accountNo;
                lineNo += 10000;
                var journaltemplate;
                var accountType;
                var docType;
                var vatAmount = data.data[i]["VATAmount"];
                var fxRate = data.data[i]["FxRate"];
                var amountLCY = vatAmount * fxRate;
                var billNo = data.data[i]["BillNo"].length;
                var invoiceSourceType;

                if (billNo == 16) {
                    invoiceSourceType = "Electronic";
                }
                else {
                    invoiceSourceType = "Paper";
                }


                var currencyCode;

                if (data.data[i]["CurrencyType"] == "TRY") {
                    currencyCode = " ";
                }
                else {
                    currencyCode = data.data[i]["CurrencyType"];
                }

                if (data.data[i]['DocumentType'] == "DocumentType02") {
                    journaltemplate = '1-INVOICE';
                    docType = 'INVOICE';
                }
                else {
                    journaltemplate = '1-MASRAF';
                    docType = 'Payment';

                }

                if (data.data[i]["VAT"] == "VAT01") {
                    accountNo = "191.01.003"
                }
                else if (data.data[i]["VAT"] == "VAT02") {
                    accountNo = "191.01.002"

                }
                else if (data.data[i]["VAT"] == "VAT03") {
                    accountNo = "191.01.001"

                }

                if (data.data[i]["CurrencyType"] != "TRY") {
                    currencyFactor = 1 / data.data[i]["FxRate"];

                }
                else {
                    currencyFactor = 0;
                }




                worksheet.addRow({ journalTemplateName: journaltemplate, journalBatchName: docNo, lineNo: lineNo, accountType: 'GL/Account', accountNo: accountNo, documentDate: docDate, postingDate: postDate, documentType: docType, documentNo: 'YDMH-' + docNo, description: description, currencyCode: currencyCode, amount: vatAmount, amountLCY: amountLCY, currencyFactor: currencyFactor, externalDocumentNo: docNo, invoiceSourceType: invoiceSourceType });

            }
            else if (data.data[i]['PaymentType'] == "PaymentType01") {
                //worksheet2.addRow({ UserName: data.data[i].PaymentType, transactionCodeI: data.data[i].PaymentType });
                var docNo = data.data[i]["DocumentNo"];
                var docDate = data.data[i]["DocumentDate"];
                var postDate = data.data[i]["Date"];
                var description = data.data[i]["Description"];
                var accountNo;
                lineNo += 10000;
                var journaltemplate;
                var accountType;
                var docType;
                var vatAmount = data.data[i]["VATAmount"];
                var fxRate = data.data[i]["FxRate"];
                var amountLCY = vatAmount * fxRate;
                var billNo = data.data[i]["BillNo"].length;
                var invoiceSourceType;

                if (billNo == 16) {
                    invoiceSourceType = "Electronic";
                }
                else {
                    invoiceSourceType = "Paper";
                }


                var currencyCode;

                if (data.data[i]["CurrencyType"] == "TRY") {
                    currencyCode = " ";
                }
                else {
                    currencyCode = data.data[i]["CurrencyType"];
                }

                if (data.data[i]['DocumentType'] == "DocumentType02") {
                    journaltemplate = '1-INVOICE';
                    docType = 'INVOICE';
                }
                else {
                    journaltemplate = '1-MASRAF';
                    docType = 'Payment';

                }

                if (data.data[i]["VAT"] == "VAT01") {
                    accountNo = "191.01.003"
                }
                else if (data.data[i]["VAT"] == "VAT02") {
                    accountNo = "191.01.002"

                }
                else if (data.data[i]["VAT"] == "VAT03") {
                    accountNo = "191.01.001"

                }

                if (data.data[i]["CurrencyType"] != "TRY") {
                    currencyFactor = 1 / data.data[i]["FxRate"];

                }
                else {
                    currencyFactor = 0;
                }




                worksheet2.addRow({ journalTemplateName: journaltemplate, journalBatchName: docNo, lineNo: lineNo, accountType: 'GL/Account', accountNo: accountNo, documentDate: docDate, postingDate: postDate, documentType: docType, documentNo: 'YDMH-' + docNo, description: description, currencyCode: currencyCode, amount: vatAmount, amountLCY: amountLCY, currencyFactor: currencyFactor, externalDocumentNo: docNo, invoiceSourceType: invoiceSourceType });

            }

        }


        var amountLCY2=0;
        var fxRate2=0;
        var total = 0;
        var totalamountlcy = 0;
        var expenseAmount2=0;
        var array = [];
        for (var k = 0; k < data.data.length; k++) {
            array.push(data.data[k]["DocumentNo"]);
        }
        var uniqueDocNo = [...new Set(array)]; //docno a göre uniq listem
        console.log(uniqueDocNo, "liste");

        for (var i = 0; i < uniqueDocNo.length; i++) {

            lineNo += 10000;
            var docNo = uniqueDocNo[i];
            var journaltemplate;
            var accountType;
            var accountNo;
            var docDate;

            var postDate;
            var docType;
            var description;
            var currencyCode;
            var vatAmount;

            var invoiceSourceType;
            var currencyFactor;
            var paymentType2;
            for (var j = 0; j < data.data.length; j++) {
                debugger;
                if (data.data[j]["DocumentNo"] == docNo) {
                    fxRate2 = data.data[j]["FxRate"];
                    var total = data.data[j]["ExpenseAmount"] + data.data[j]["VATAmount"];
                    var totalamountlcy = data.data[j]["AmountLCY"];
                    amountLCY2 += total * fxRate2;
                    expenseAmount2 +=total;
                    currencyCode = data.data[j]["CurrencyType"];

                    docDate = data.data[j]["DocumentDate"];
                    postDate = data.data[j]["Date"];
                    if (data.data[j]["DocumentType"] == "DocumentType02") {
                        journaltemplate = "1-INVOICE";
                        accountNo = data.data[j]["VendorNo"];
                        docType = "INVOICE"
                    }
                    else {
                        journaltemplate = "1-MASRAF";
                    }
                    if (data.data[j]["CurrencyType"]!="TRY") {
                        currencyFactor = 1 / data.data[j]["FxRate"];

                    }
                    else {
                        currencyFactor = 0;
                    }
                    if (data.data[j]["BillNo"].length == 16) {
                        invoiceSourceType = "Electronic";
                    }
                    else {
                        invoiceSourceType = "Paper";
                    }
                    if (accountNo.startsWith('300') || accountNo.startsWith('320')) {
                        accountType = 'Vendor';
                    }
                    else if (accountNo.startsWith('120') || accountNo.startsWith('195')) {
                        accountType = 'Customer';

                    }
                    else {
                        accountType = 'GL/Account';

                    }
                    paymentType2 = data.data[j]["PaymentType"];
                }

            }
            if (paymentType2 == "PaymentType01") {
                worksheet2.addRow({ journalTemplateName: journaltemplate, journalBatchName: docNo, lineNo: lineNo, accountType: accountType, accountNo: accountNo, documentDate: docDate, postingDate: postDate, documentType: docType, documentNo: 'YDMH-' + docNo, description: docNo, currencyCode: currencyCode, amount: -expenseAmount2, amountLCY: -amountLCY2, currencyFactor: currencyFactor, externalDocumentNo: docNo, invoiceSourceType: invoiceSourceType });
                total = 0;
                expenseAmount2 = 0;
                amountLCY2 = 0;
                totalamountlcy = 0;
                fxRate2 = 0;
            }
            else if (paymentType2 == "PaymentType02") {
                worksheet.addRow({ journalTemplateName: journaltemplate, journalBatchName: docNo, lineNo: lineNo, accountType: accountType, accountNo: accountNo, documentDate: docDate, postingDate: postDate, documentType: docType, documentNo: 'YDMH-' + docNo, description: docNo, currencyCode: currencyCode, amount: -expenseAmount2, amountLCY: -amountLCY2, currencyFactor: currencyFactor, externalDocumentNo: docNo, invoiceSourceType: invoiceSourceType });
                total = 0;
                expenseAmount2 = 0;
                amountLCY2 = 0;
                totalamountlcy = 0;
                fxRate2 = 0;
            }


        }



        workbook.xlsx.writeBuffer().then(function (buffer) {
        saveAs(new Blob([buffer], { type: 'application/octet-stream' }), '@DBResources.GetText("BC_Export")'+'.xlsx');
        });

    }
