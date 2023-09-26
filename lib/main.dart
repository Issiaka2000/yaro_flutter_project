import 'package:excel/excel.dart';
import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart';
import 'dart:io';
import 'dart:typed_data';

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  // This widget is the root of your application.
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Flutter Demo',
      debugShowCheckedModeBanner: false,
      theme: ThemeData(
        primarySwatch: Colors.blue,
      ),
      home: const MyHomePage(title: 'Flutter Demo Home Page'),
    );
  }
}

class MyHomePage extends StatefulWidget {
  const MyHomePage({super.key, required this.title});
  final String title;
  @override
  State<MyHomePage> createState() => _MyHomePageState();
}

class _MyHomePageState extends State<MyHomePage> {
  @override
  Widget build(BuildContext context) {
    var matieresParJour = {
     'data': {'Monday': [
        {'matiere': 'Mathematics', 'heure': '08.00 - 09.30'},
        {'matiere': 'Physics', 'heure': '09.45 - 11.15'},
        {'matiere': 'French', 'heure': '11.30 - 13.00'},
        {'matiere': 'English', 'heure': '14.00 - 15.30'},
        {'matiere': 'History', 'heure': '15.45 - 17.15'},
        {'matiere': 'PE', 'heure': '17.30 - 18.00'},
      ],
      'Tuesday': [
        {'matiere': 'French', 'heure': '08.00 - 09.30'},
        {'matiere': 'Physics', 'heure': '09.45 - 11.15'},
        {'matiere': 'Mathematics', 'heure': '11.30 - 13.00'},
        {'matiere': 'Biology', 'heure': '14.00 - 15.30'},
        {'matiere': 'PE', 'heure': '15.45 - 17.15'},
        {'matiere': 'English', 'heure': '17.30 - 18.00'},
      ],
      'Wednesday': [
        {'matiere': 'English', 'heure': '08.00 - 09.30'},
        {'matiere': ' ', 'heure': '09.45 - 11.15'},
        {'matiere': 'History', 'heure': '11.30 - 13.00'},
        {'matiere': 'French', 'heure': '14.00 - 15.30'},
        {'matiere': 'Mathematics', 'heure': '15.45 - 17.15'},
        {'matiere': 'Physics', 'heure': '17.30 - 18.00'},
      ],
      'Thursday': [
        {'matiere': 'PE', 'heure': '08.00 - 09.30'},
        {'matiere': 'Mathematics', 'heure': '09.45 - 11.15'},
        {'matiere': 'English', 'heure': '11.30 - 13.00'},
        {'matiere': 'Physics', 'heure': '14.00 - 15.30'},
        {'matiere': 'Biology', 'heure': '15.45 - 17.15'},
        {'matiere': 'French', 'heure': '17.30 - 18.00'},
      ],
      'Friday': [
        {'matiere': 'Physics', 'heure': '08.00 - 09.30'},
        {'matiere': 'French', 'heure': '09.45 - 11.15'},
        {'matiere': 'English', 'heure': '11.30 - 13.00'},
        {'matiere': 'PE', 'heure': '14.00 - 15.30'},
        {'matiere': 'Mathematics', 'heure': '15.45 - 17.15'},
        {'matiere': 'History', 'heure': '17.30 - 18.00'},
      ],
      'Saturday': [
        {'matiere': 'French', 'heure': '08.00 - 09.30'},
        {'matiere': 'Mathematics', 'heure': '09.45 - 11.15'},
        {'matiere': 'English', 'heure': '11.30 - 13.00'},
        {'matiere': 'PE', 'heure': '14.00 - 15.30'},
        {'matiere': 'Biology', 'heure': '15.45 - 17.15'},
        {'matiere': 'Physics', 'heure': '17.30 - 18.00'},
      ]},
      'metadata':{
        'week':5,
        'date':'17/05/2023',
        'classroom':'Amphi, 100',
        'filiere':'Electrical Ingeniering 24',
      }
    };

    return Scaffold(
      appBar: AppBar(
        title: const Text('Excel file template'),
      ),
      body: Padding(
        padding: const EdgeInsets.all(8.0),
        child: SizedBox(
          width: double.infinity,
          child: TextButton.icon(
              onPressed: () {
                generateCustomExcel(matieresParJour);
              },
              icon: const Icon(Icons.generating_tokens),
              label: const Text("Creer un model")),
        ),
      ),
    );
  }

  Future<void> generateCustomExcel(
      Map<String, Map<dynamic, dynamic>> matieresParJour) async {
    CellStyle cellstyle(bool bold, String color, int fontsize) {
      return CellStyle(
        bold: bold,
        textWrapping: TextWrapping.WrapText,
        horizontalAlign: HorizontalAlign.Center,
        verticalAlign: VerticalAlign.Center,
        backgroundColorHex: color,
        fontFamily: getFontFamily(FontFamily.Calibri),
        fontSize: fontsize,
      );
    }

    final excel = Excel.createExcel();
    final sheet = excel.sheets[excel.getDefaultSheet()];

    // fusion des cellules
    sheet!.merge(CellIndex.indexByString('B1'), CellIndex.indexByString('B5'),
        customValue: '');
    sheet.merge(CellIndex.indexByString('C6'), CellIndex.indexByString('H6'),
        customValue: 'Week');
    sheet.merge(CellIndex.indexByString('D7'), CellIndex.indexByString('F7'),
        customValue: 'Week');
    sheet.merge(CellIndex.indexByString('D9'), CellIndex.indexByString('F9'),
        customValue: 'Week');
    sheet.merge(CellIndex.indexByString('C12'), CellIndex.indexByString('H12'),
        customValue: 'break');
    sheet.merge(CellIndex.indexByString('C14'), CellIndex.indexByString('D14'),
        customValue: 'break');
    sheet.merge(CellIndex.indexByString('F14'), CellIndex.indexByString('H14'),
        customValue: 'break');
    sheet.merge(CellIndex.indexByString('C16'), CellIndex.indexByString('H16'),
        customValue: 'break');
    sheet.merge(CellIndex.indexByString('C18'), CellIndex.indexByString('H18'),
        customValue: 'break');
    sheet.merge(CellIndex.indexByString('C20'), CellIndex.indexByString('H20'),
        customValue: 'break');
    sheet.merge(CellIndex.indexByString('E13'), CellIndex.indexByString('E14'),
        customValue: 'break');

// for the line 5
    final cellStyle1 = cellstyle(true, '#E8D519', 13);
    final weekcellStyle = cellstyle(true, 'none', 13);
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: 5))
      ..value = matieresParJour['metadata']!['filiere']
      ..cellStyle = cellStyle1;
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: 5))
      ..value = 'Week '+matieresParJour['metadata']!['week'].toString()
      ..cellStyle = weekcellStyle;

// for the line 8
    final cellStyleB8 = cellstyle(true, '#E7E3BD', 10);
    final cellStyleC8 = cellstyle(true, '#E8D519', 10);
    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: 7))
      ..value = 'Enter Start Date'
      ..cellStyle = cellStyleB8;

    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: 7))
      ..value = matieresParJour['metadata']!['date'] 
      ..cellStyle = cellStyleC8;

    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 4, rowIndex: 7))
      ..value = 'Classroom'
      ..cellStyle = cellStyleB8;

    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 5, rowIndex: 7))
      ..value = matieresParJour['metadata']!['classroom']
      ..cellStyle = cellStyleC8;

// to manage the house column
    List hourse = [
      '08.00 - 09.30',
      '09.30 - 09.45',
      '09.45 - 11.15',
      '11.15 - 11.30',
      '11.30 - 13.00',
      '13.00 - 14.00',
      '14.00 - 15.30',
      '15.30 - 15.45',
      '15.45 - 17.15',
      '17.15 - 17.30',
      '17.30 - 18.00'
    ];
    final cellStylehouse = cellstyle(false, 'none', 09);
    for (var row = 11; row < 22; row++) {
      sheet.cell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: row - 1))
        ..value = hourse[row - 11]
        ..cellStyle = cellStylehouse;
    }

    // Créer un style de cellule pour la première ligne
    final headerscellStyle = cellstyle(true, '#E7E3BD', 10);
    // Appliquer le style à la première ligne
    List days = [
      'Schedule',
      'Monday',
      'Tuesday',
      'Wednesday',
      'Thursday',
      'Friday',
      'Saturday'
    ];
    for (var col = 1; col < 8; col++) {
      sheet.cell(CellIndex.indexByColumnRow(columnIndex: col, rowIndex: 9))
        ..value = days[col - 1]
        ..cellStyle = headerscellStyle;
    }

    // Remplir les cellules des colonnes de chaque jour avec les matières
    for (var col = 2; col < 8; col++) {
      final jour = days[col - 1];
      final matieres = matieresParJour['data']![jour];

      if (matieres != null) {
        for (var i = 0; i < matieres.length; i++) {
          final row = 10 + (i * 2);
          final cellValue = matieres[i]['matiere'];
          final curenthouseValue = matieres[i]['heure'];
          final presentheurevalue = sheet
              .cell(CellIndex.indexByColumnRow(columnIndex: 1, rowIndex: row))
              .value.toString();
        presentheurevalue==curenthouseValue?
          (sheet
              .cell(CellIndex.indexByColumnRow(columnIndex: col, rowIndex: row))
            ..value = cellValue
            ..cellStyle = cellStylehouse):(sheet
              .cell(CellIndex.indexByColumnRow(columnIndex: col, rowIndex: row))
            ..value = '');

     
        }
      }
    }

    //Style for merge cells
    List merges = [11, 17, 19];
    final breakcellStyle = cellstyle(true, '#DEF9C3', 09);
    final lunchcellStyle = cellstyle(true, '#B5DC8E', 09);
    for (var i = 0; i < merges.length; i++) {
      sheet
          .cell(CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: merges[i]))
        ..value = 'Break'
        ..cellStyle = breakcellStyle;
    }

    List merge2 = [2, 5];
    for (var i = 0; i < merge2.length; i++) {
      sheet.cell(
          CellIndex.indexByColumnRow(columnIndex: merge2[i], rowIndex: 13))
        ..value = 'Break'
        ..cellStyle = breakcellStyle;
    }

    sheet.cell(CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: 15))
      ..value = 'Lunch'
      ..cellStyle = lunchcellStyle;

    // wooman ampowerment cell
    final empowermentcellStyle = cellstyle(false, '#E8D519', 09);
    sheet.cell(CellIndex.indexByString('E13'))
      ..value = 'Female Empowerment (10h-11h30)'
      ..cellStyle = empowermentcellStyle;
//
    final File file = File("custom_excel.xlsx");
    file.createSync(recursive: true);
    // Écrivez le contenu Excel dans le fichier
    file.writeAsBytesSync(excel.save()!);
    print('Fichier Excel personnalisé créé avec succès.');
  }
}
