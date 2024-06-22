import 'dart:io';
import 'package:flutter/material.dart';
import 'package:intl/intl.dart';
import 'package:path_provider/path_provider.dart';
import 'package:open_file/open_file.dart';

void main() {
  runApp(MyApp());
}

class InvoiceData {
  String invoiceNumber = '';
  String date = '';
  String service = '';
  String salary1 = '';
  String salary2 = '';
  String salary3 = '';
}

class MyApp extends StatelessWidget {
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      home: Scaffold(
        appBar: AppBar(title: Text('Rechnungserstellung')),
        body: Padding(
          padding: EdgeInsets.all(10),
          child: InvoiceForm(),
        ),
      ),
    );
  }
}

class InvoiceForm extends StatefulWidget {
  @override
  _InvoiceFormState createState() => _InvoiceFormState();
}

class _InvoiceFormState extends State<InvoiceForm> {
  final InvoiceData invoiceData = InvoiceData();
  final GlobalKey<FormState> _formKey = GlobalKey<FormState>();

  @override
  Widget build(BuildContext context) {
    return Form(
      key: _formKey,
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: <Widget>[
          TextFormField(
            decoration: InputDecoration(labelText: 'Rechnungsnummer'),
            onChanged: (value) => invoiceData.invoiceNumber = value,
            validator: (value) {
              if (value!.isEmpty) {
                return 'Rechnungsnummer eingeben';
              }
              return null;
            },
          ),
          TextFormField(
            decoration: InputDecoration(labelText: 'Rechnungsdatum (TT.MM.JJJJ)'),
            onChanged: (value) => invoiceData.date = value,
            validator: (value) {
              if (value!.isEmpty) {
                return 'Rechnungsdatum eingeben';
              }
              // Additional date validation logic can be added here
              return null;
            },
          ),
          TextFormField(
            decoration: InputDecoration(labelText: 'Leistungszeitraum'),
            onChanged: (value) => invoiceData.service = value,
            validator: (value) {
              if (value!.isEmpty) {
                return 'Leistungszeitraum eingeben';
              }
              return null;
            },
          ),
          TextFormField(
            decoration: InputDecoration(labelText: 'Commessa CM0184-003 Assistenza di Commessa Costr 720'),
            onChanged: (value) => invoiceData.salary1 = value,
            validator: (value) {
              if (value!.isEmpty) {
                return 'Gehalt eingeben';
              }
              return null;
            },
          ),
          TextFormField(
            decoration: InputDecoration(labelText: 'Commessa CM0189-003 Assistenza di Commessa Costr 718'),
            onChanged: (value) => invoiceData.salary2 = value,
            validator: (value) {
              if (value!.isEmpty) {
                return 'Gehalt eingeben';
              }
              return null;
            },
          ),
          TextFormField(
            decoration: InputDecoration(labelText: 'Commessa CM0231-003 Assistenza di Commessa Costr 721'),
            onChanged: (value) => invoiceData.salary3 = value,
            validator: (value) {
              if (value!.isEmpty) {
                return 'Gehalt eingeben';
              }
              return null;
            },
          ),
          ElevatedButton(
            onPressed: () {
              if (_formKey.currentState!.validate()) {
                createInvoice();
              }
            },
            child: Text('Rechnung erstellen'),
          ),
        ],
      ),
    );
  }

  Future<void> createInvoice() async {
    // Check if all fields are valid
    if (_formKey.currentState!.validate()) {
      // Parse salaries to doubles
      double salary1, salary2, salary3;
      try {
        salary1 = double.parse(invoiceData.salary1);
        salary2 = double.parse(invoiceData.salary2);
        salary3 = double.parse(invoiceData.salary3);
      } catch (e) {
        showErrorDialog('Bitte geben Sie gültige Zahlen für die Gehälter ein');
        return;
      }

      // Validate and format date
      DateTime parsedDate;
      try {
        parsedDate = DateFormat('dd.MM.yyyy').parse(invoiceData.date);
      } catch (e) {
        showErrorDialog('Ungültiges Datum. Bitte im Format TT.MM.JJJJ eingeben.');
        return;
      }

      // Prepare context data
      Map<String, dynamic> contextData = {
        'DATE': DateFormat('dd.MM.yyyy').format(parsedDate),
        'INVOICE_NUMBER': invoiceData.invoiceNumber,
        'SERVICE': invoiceData.service,
        'SALARY1': '${salary1.toStringAsFixed(2)}',
        'SALARY2': '${salary2.toStringAsFixed(2)}',
        'SALARY3': '${salary3.toStringAsFixed(2)}',
        'BRUTTO': '${(salary1 + salary2 + salary3).toStringAsFixed(2)}',
        'STEUER': '${((salary1 + salary2 + salary3) * 0.19).toStringAsFixed(2)}',
        'NETTO': '${((salary1 + salary2 + salary3) * 0.81).toStringAsFixed(2)}',
      };

      // Create a temporary directory to store the generated DOCX
      final tempDir = await getTemporaryDirectory();
      final tempPath = '${tempDir.path}/Rechnung_Nr${invoiceData.invoiceNumber}.txt';

      // Write invoice data to a text file
      File tempFile = File(tempPath);
      await tempFile.writeAsString(_buildInvoiceText(contextData));

      // Show success message and open the file
      showDialog(
        context: context,
        builder: (BuildContext context) => AlertDialog(
          title: Text('Erfolg'),
          content: Text('Rechnung wurde erfolgreich erstellt und gespeichert!'),
          actions: <Widget>[
            TextButton(
              child: Text('OK'),
              onPressed: () {
                Navigator.of(context).pop();
                OpenFile.open(tempPath);
              },
            ),
          ],
        ),
      );
    }
  }

  String _buildInvoiceText(Map<String, dynamic> data) {
    return '''
    Rechnungsdatum: ${data['DATE']}
    Rechnungsnummer: ${data['INVOICE_NUMBER']}
    Leistungszeitraum: ${data['SERVICE']}
    Gehalt 1: ${data['SALARY1']}
    Gehalt 2: ${data['SALARY2']}
    Gehalt 3: ${data['SALARY3']}
    BRUTTO: ${data['BRUTTO']}
    STEUER: ${data['STEUER']}
    NETTO: ${data['NETTO']}
    ''';
  }

  void showErrorDialog(String message) {
    showDialog(
      context: context,
      builder: (BuildContext context) => AlertDialog(
        title: Text('Fehler'),
        content: Text(message),
        actions: <Widget>[
          TextButton(
            child: Text('OK'),
            onPressed: () => Navigator.of(context).pop(),
          ),
        ],
      ),
    );
  }
}
