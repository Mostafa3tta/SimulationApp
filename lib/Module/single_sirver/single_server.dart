import 'dart:convert';
import 'dart:io';
import 'package:flutter/material.dart';
import 'package:hexcolor/hexcolor.dart';
import 'package:open_file/open_file.dart';
import 'package:path_provider/path_provider.dart';
import 'package:syncfusion_flutter_xlsio/xlsio.dart';
import 'package:universal_html/html.dart' show AnchorElement;
import 'package:flutter/foundation.dart' show kIsWeb;

class homePageScreen extends StatelessWidget {
  int number;

  homePageScreen(this.number);

  List<String> names = [
    'Janelle Clark',
    'Cristina Mccaffrey',
    'Dewi Duran',
    'Molly Coates',
    'Isabelle Witt',
    'Brenna Morgan',
    'Ethan Harrington',
    'Johnathan French',
    'Jac Farrell',
    'Shane Thomas',
    'Elspeth Derrick',
    'Teegan Mackay',
    'Mylah Woodley',
    'Alec Iles',
    'Thea Cousins',
    'Madeleine Clements',
    'Zakariya Millar',
    'Johanna Hogg',
    'Kayson Marquez',
    'Kelsey Thornton',
    'Ariyan Goff',
    'Faith Franklin',
    'Raheel Bradshaw',
    'Zahra Wood',
    'Susan Lu',
    'Troy Duarte',
    "Samiyah Joyce",
    'Rumaisa Robin',
    "Marguerite Robertson",
    'Alexia Draper',
    'Essa Harvey',
    'Callum Wormald',
    'Kairo Lott',
    'Willow Hutchings',
    'Cairon Clarkson',
    'Jai Roach',
    'Ilayda Hulme',
    'Elowen Hanna',
    'Tiah Nielsen',
    'Ziva Pineda',
    'Abubakr Mcintosh',
    'Hakeem Gaines',
    'Maureen Elliott',
    'Christopher Byers',
    'Maryam Larsen',
    'Ritik Stout',
    'Hashim Montgomery',
    'Ayla Dougherty',
    "Tyler-James Calderon",
    'Robson Pike',
  ];
  List arrivalBetween = [1, 2, 3, 4, 5, 6, 7, 8];
  List propability = [
    0.125,
    0.125,
    0.125,
    0.125,
    0.125,
    0.125,
    0.125,
    0.125,
  ];
  List ser_time = [1, 2, 3, 4, 5, 6];
  List ser_propability = [0.1, 0.2, 0.3, 0.25, 0.1, 0.05];
  Map service = {
    1: 'Open',
    2: 'Delete',
    3: 'Deposite',
    4: 'Withdraw',
    5: 'Transfer',
    6: 'Inquiry',
    7: 'Report',
  };
  Map duration = {
    'Open': 10,
    'Delete': 15,
    'Deposite': 5,
    'Withdraw': 7,
    'Transfer': 8,
    'Inquiry': 3,
    'Report': 4,
  };

  @override
  Widget build(BuildContext context) {
    return Container(
        decoration: BoxDecoration(
            gradient: LinearGradient(colors: [
          HexColor("#acb6e5"),
          HexColor("#86fde8"),
        ])),
        child: Scaffold(
          backgroundColor: Colors.transparent,
          body: Container(
            child: Center(
              child: ElevatedButton(
                child: const Text("Create Exel"),
                onPressed: createExel,
              ),
            ),
          ),
        ));
  }

  Future<void> createExel() async {
    final Workbook workbook = Workbook();
    final Worksheet sheet = workbook.worksheets[0];
    final Worksheet sheet2 = workbook.worksheets.add();

    void sheet_1(){

      void arrivalPropability() {
        sheet.getRangeByName("A1:Q10").cellStyle.hAlign = HAlignType.center;
        sheet.getRangeByName("A1").setText("Time Between Arrivals (Minutes)");
        sheet.getRangeByName("A1:E9").cellStyle.borders.all.lineStyle = LineStyle.thin;
        sheet.getRangeByName("A1:E1").cellStyle.backColorRgb = Colors.green;
        sheet.getRangeByName("A1:E1").cellStyle.borders.all.lineStyle =
            LineStyle.thin;
        sheet.getRangeByName("A1").columnWidth = 28;

        for (int i = 0; i < arrivalBetween.length; i++) {
          sheet.getRangeByName("A${i + 2}").setText('${arrivalBetween[i]}');
        }
        sheet.getRangeByName("B1").setText("Propability");
        sheet.getRangeByName("B1").columnWidth = 12;

        for (int i = 0; i < propability.length; i++) {
          sheet.getRangeByName("B${i + 2}").setNumber(propability[i]);
        }
        sheet.getRangeByName("C1").setText("Cumulative");
        sheet.getRangeByName("C1").columnWidth = 10;

        for (int i = 0; i < propability.length - 1; i++) {
          sheet.getRangeByName("C2").setFormula("${propability[0]}");
          sheet
              .getRangeByName("C${i + 3}")
              .setFormula("=SUM(${propability[0]}+${"C${i + 2}"})");
        }
        sheet.getRangeByName("E1").setText("To");

        for (int i = 0; i < propability.length; i++) {
          sheet.getRangeByName("E${i + 2}").setFormula("=C${i + 2} * 1000");
        }
        sheet.getRangeByName("D1").setText("From");

        for (int i = 0; i < propability.length - 1; i++) {
          sheet.getRangeByName("D2").setFormula("=0");
          sheet.getRangeByName("D${i + 3}").setFormula("=SUM(E${i + 2} + 1)");
        }
        sheet.getRangeByName("A10:E10").merge();
        sheet.getRangeByName("A10:E10").setText("Arrival Propability");

        sheet.getRangeByName("A10:E10").cellStyle.backColorRgb =
            Colors.amberAccent;
        sheet.getRangeByName("A10:E10").cellStyle.borders.all.lineStyle =
            LineStyle.medium;
      }

      void servicesPropability() {
        sheet.getRangeByName("H1").setText("Service Time (Minutes)");
        sheet.getRangeByName("H1:L9").cellStyle.borders.all.lineStyle = LineStyle.thin;
        sheet.getRangeByName("H1:L1").cellStyle.backColorRgb = Colors.green;
        sheet.getRangeByName("H1:L1").cellStyle.borders.all.lineStyle =
            LineStyle.thin;
        sheet.getRangeByName("H1").columnWidth = 19;
        for (int i = 0; i < ser_time.length; i++) {
          sheet.getRangeByName("H${i + 2}").setText('${ser_time[i]}');
        }
        sheet.getRangeByName("I1").setText("Propability");
        sheet.getRangeByName("I1").columnWidth = 12;
        for (int i = 0; i < ser_propability.length; i++) {
          sheet.getRangeByName("I${i + 2}").setNumber(ser_propability[i]);
        }
        sheet.getRangeByName("J1").setText("Cumulative");
        sheet.getRangeByName("J1").columnWidth = 10;
        for (int i = 0; i < ser_propability.length - 1; i++) {
          sheet.getRangeByName("J2").setFormula("${ser_propability[0]}");
          sheet
              .getRangeByName("J${i + 3}")
              .setFormula("=SUM(${"I${i + 3}"}+${"J${i + 2}"})");
        }
        sheet.getRangeByName("L1").setText("To");
        for (int i = 0; i < ser_propability.length; i++) {
          sheet.getRangeByName("L${i + 2}").setFormula("=J${i + 2} * 100");
        }
        sheet.getRangeByName("K1").setText("From");
        for (int i = 0; i < ser_propability.length - 1; i++) {
          sheet.getRangeByName("K2").setFormula("=0");
          sheet.getRangeByName("K${i + 3}").setFormula("=SUM(L${i + 2} + 1)");
        }
        sheet.getRangeByName("H10:L10").merge();
        sheet.getRangeByName("H10:L10").setText("Services Propability");
        sheet.getRangeByName("H10:L10").cellStyle.backColorRgb =
            Colors.amberAccent;
        sheet.getRangeByName("H10:L10").cellStyle.borders.all.lineStyle =
            LineStyle.medium;
      }

      void Services() {
        sheet.getRangeByName("O1").setText("Service ID");
        sheet.getRangeByName("O1:Q9").cellStyle.borders.all.lineStyle = LineStyle.thin;
        sheet.getRangeByName("O1:Q1").cellStyle.backColorRgb = Colors.green;
        sheet.getRangeByName("O1:Q1").cellStyle.borders.all.lineStyle =
            LineStyle.thin;
        sheet.getRangeByName("O1").columnWidth = 11;

        for (int i = 1; i <= service.length; i++) {
          sheet.getRangeByName("O${i + 1}").setFormula('${i}');
        }
        sheet.getRangeByName("P1").setText("Service");
        sheet.getRangeByName("P1").columnWidth = 9;

        for (int i = 0; i < service.length; i++) {
          sheet
              .getRangeByName("P${i + 2}")
              .setText('${service.values.elementAt(i)}');
        }
        sheet.getRangeByName("Q1").setText("Service Duration");
        sheet.getRangeByName("Q1").columnWidth = 17;

        for (int i = 0; i < duration.length; i++) {
          sheet
              .getRangeByName("Q${i + 2}")
              .setText('${duration.values.elementAt(i)}');
        }
        sheet.getRangeByName("O10:Q10").merge();
        sheet.getRangeByName("O10:Q10").setText("Services");

        sheet.getRangeByName("O10:Q10").cellStyle.backColorRgb =
            Colors.amberAccent;
        sheet.getRangeByName("O10:Q10").cellStyle.borders.all.lineStyle =
            LineStyle.medium;
      }

      arrivalPropability();
      servicesPropability();
      Services();
    }
    void sheet_2(){

      void simulationTable() {
        sheet2.getRangeByName("A3:M3").cellStyle.backColorRgb = Colors.green;
        for (int i = 1; i <= number; i++) {
          sheet2
              .getRangeByName("A${i + 2}:M${number + 3}")
              .cellStyle
              .borders
              .all
              .lineStyle = LineStyle.thin;
        }
        sheet2.getRangeByName("A1:M${number + 3}").cellStyle.hAlign =
            HAlignType.center;
        sheet2.getRangeByName("F2:H2").merge();
        sheet2.getRangeByName("F2:H2").setText("Queuing System");
        sheet2.getRangeByName("F2:H2").cellStyle.borders.all.lineStyle =
            LineStyle.medium;
        sheet2.getRangeByName("F2:H2").cellStyle.backColorRgb =
            Colors.amberAccent;

        sheet2.getRangeByName("A3").setText("Customer (Number)");
        sheet2.getRangeByName("A3").columnWidth = 16;
        for (int i = 1; i <= number; i++) {
          sheet2.getRangeByName("A${i + 3}").setText('${i}');
        }
        sheet2.getRangeByName("B3").setText("Random.Digit"); // =INT(RAND()*1000)
        sheet2.getRangeByName("B3").columnWidth = 11;
        sheet2.getRangeByName("B4").setNumber(0);
        for (int i = 1; i <= number - 1; i++) {
          sheet2.getRangeByName("B${i + 4}").setFormula("=INT(RAND()*1000)");
        }
        sheet2.getRangeByName("C3").setText("inter arrival (Time)");
        sheet2.getRangeByName("C3").columnWidth = 16;
        sheet2.getRangeByName("C4").setNumber(0);
        for (int i = 1; i <= number - 1; i++) {
          sheet2.getRangeByName("C${i + 4}").setFormula(
              "=LOOKUP(Sheet2!B${i + 4},Sheet1!D${2}:E${9},Sheet1!A${2}:A${9})");
        }
        sheet2.getRangeByName("D3").setText("Arrival Time (Clock)");
        sheet2.getRangeByName("D3").columnWidth = 16;
        sheet2.getRangeByName("D4").setNumber(0);
        for (int i = 1; i <= number - 1; i++) {
          sheet2.getRangeByName("D${i + 4}").setFormula("==D${i + 3}+C${i + 4}");
        }
        sheet2.getRangeByName("E3").setText("Service Code");
        sheet2.getRangeByName("E3").columnWidth = 11;
        for (int i = 1; i <= number; i++) {
          sheet2.getRangeByName("E${i + 3}").setFormula("=INT(RAND()*7)+1");
        }
        sheet2.getRangeByName("F3").setText("Service Title");
        sheet2.getRangeByName("F3").columnWidth = 10;
        for (int i = 1; i <= number; i++) {
          sheet2.getRangeByName("F${i + 3}").setFormula(
              "=LOOKUP(E${i + 3},Sheet1!O${2}:O${8},Sheet1!P${2}:P${8})");
        }
        sheet2.getRangeByName("G3").setText("Time Service Begins (Clock)");
        sheet2.getRangeByName("G3").columnWidth = 22;
        sheet2.getRangeByName("G4").setNumber(0);
        for (int i = 1; i <= number - 1; i++) {
          sheet2
              .getRangeByName("G${i + 4}")
              .setFormula("=MAX(J${i + 3},D${i + 4})");
        }
        sheet2.getRangeByName("H3").setText("Random.Digit");
        sheet2.getRangeByName("H3").columnWidth = 11;
        sheet2.getRangeByName("H4").setNumber(0);
        for (int i = 1; i <= number - 1; i++) {
          sheet2.getRangeByName("H${i + 4}").setFormula("=INT(RAND()*100)");
        }
        sheet2.getRangeByName("I3").setText("Service Time (Duration)");
        sheet2.getRangeByName("I3").columnWidth = 19;
        for (int i = 1; i <= number; i++) {
          sheet2.getRangeByName("I${i + 3}").setFormula(
              "==LOOKUP(H${i + 3},Sheet1!K${2}:L${7},Sheet1!H${2}:H${7})");
        }
        sheet2.getRangeByName("J3").setText("Time Service Ends (Clock)");
        sheet2.getRangeByName("J3").columnWidth = 21;
        for (int i = 1; i <= number; i++) {
          sheet2
              .getRangeByName("J${i + 3}")
              .setFormula("=SUM(G${i + 3}+I${i + 3})");
        }
        sheet2.getRangeByName("K3").setText("Waiting Time");
        sheet2.getRangeByName("K3").columnWidth = 11;
        sheet2.getRangeByName("K4").setNumber(0);
        for (int i = 1; i <= number - 1; i++) {
          sheet2.getRangeByName("K${i + 4}").setFormula("=G${i + 4}-D${i + 4}");
        }
        sheet2.getRangeByName("L3").setText("Idle Time");
        sheet2.getRangeByName("L3").columnWidth = 8;
        sheet2.getRangeByName("L4").setNumber(0);
        for (int i = 1; i <= number - 1; i++) {
          sheet2.getRangeByName("L${i + 4}").setFormula("=G${i + 4}-J${i + 3}");
        }
        sheet2.getRangeByName("M3").setText("Spent Time");
        sheet2.getRangeByName("M3").columnWidth = 9;
        for (int i = 1; i <= number; i++) {
          sheet2.getRangeByName("M${i + 3}").setFormula("=J${i + 3}-D${i + 3}");
        }
      }

      simulationTable();
    }

    sheet_1();
    sheet_2();




    final List<int> bytes = workbook.saveAsStream();
    workbook.dispose();

    if (kIsWeb) {
      AnchorElement(
          href:
              'data:application/octet-stream;charset=utf-16le;base64,${base64.encode(bytes)}')
        ..setAttribute('download', 'Queuing_System.xlsx')
        ..click();
    } else {
      final String path = (await getApplicationSupportDirectory()).path;
      final String fileName = '$path/Queuing_System.xlsx';
      final File file = File(fileName);
      await file.writeAsBytes(bytes, flush: true);
      OpenFile.open(fileName);
    }
  }
}
