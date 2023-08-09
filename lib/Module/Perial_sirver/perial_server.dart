import 'dart:io';
import 'dart:convert';
import 'package:flutter/material.dart';
import 'package:hexcolor/hexcolor.dart';
import 'package:open_file/open_file.dart';
import 'package:path_provider/path_provider.dart';
import 'package:syncfusion_flutter_xlsio/xlsio.dart';
import 'package:universal_html/html.dart' show AnchorElement;
import 'package:flutter/foundation.dart' show kIsWeb;

class Perial_Screen extends StatelessWidget {
  int number;

  Perial_Screen(this.number);

  List arrivalBetween = [1, 2, 3, 4, 5, 6];
  List propability = [0.25, 0.4, 0.2, 0.15];
  List ser_propability = [0.3, 0.28, 0.25, 0.17];
  List ser2_propability = [0.35, 0.25, 0.2, 0.2];

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

    void sheet_1() {
      void arrivalPropability() {
        void style() {
          sheet.getRangeByName("A1:R6").cellStyle.hAlign = HAlignType.center;
          sheet.getRangeByName("A1:E1").cellStyle.backColorRgb = Colors.green;
          sheet.getRangeByName("A6:E6").cellStyle.backColorRgb =
              Colors.amberAccent;
          sheet.getRangeByName("A1:E5").cellStyle.borders.all.lineStyle = LineStyle.thin;
          sheet.getRangeByName("A6:E6").cellStyle.borders.all.lineStyle = LineStyle.medium;
          sheet.getRangeByName("A1:E1").cellStyle.borders.all.lineStyle =
              LineStyle.thin;
          sheet.getRangeByName("A1").columnWidth = 28;
          sheet.getRangeByName("C1").columnWidth = 10;
          sheet.getRangeByName("B1").columnWidth = 12;
        }

        void title() {
          sheet.getRangeByName("A1").setText("Time Between Arrivals (Minutes)");
          sheet.getRangeByName("B1").setText("Propability");
          sheet.getRangeByName("C1").setText("Cumulative");
          sheet.getRangeByName("D1").setText("From");
          sheet.getRangeByName("A6:E6").merge();
          sheet.getRangeByName("A6:E6").setText("Arrival Propability");
          sheet.getRangeByName("E1").setText("To");
        }

        void process() {
          for (int i = 0; i < arrivalBetween.length - 2; i++) {
            sheet.getRangeByName("A${i + 2}").setText('${arrivalBetween[i]}');
          }
          for (int i = 0; i < propability.length; i++) {
            sheet.getRangeByName("B${i + 2}").setNumber(propability[i]);
          }

          for (int i = 0; i < propability.length - 1; i++) {
            sheet.getRangeByName("C2").setFormula("${propability[0]}");
            sheet
                .getRangeByName("C${i + 3}")
                .setFormula("=SUM(C${i + 2},B${i + 3})");
          }
          for (int i = 0; i < propability.length; i++) {
            sheet.getRangeByName("E${i + 2}").setFormula("=C${i + 2} * 100");
          }
          for (int i = 0; i < propability.length - 1; i++) {
            sheet.getRangeByName("D2").setFormula("=0");
            sheet.getRangeByName("D${i + 3}").setFormula("=SUM(E${i + 2} + 1)");
          }
        }

        title();
        process();
        style();
      }

      void abdoServer() {
        void style() {
          sheet.getRangeByName("H1:L1").cellStyle.backColorRgb = Colors.green;
          sheet.getRangeByName("H1:L1").cellStyle.borders.all.lineStyle =
              LineStyle.thin;
          sheet.getRangeByName("H1").columnWidth = 19;
          sheet.getRangeByName("I1").columnWidth = 12;
          sheet.getRangeByName("J1").columnWidth = 10;
          sheet.getRangeByName("H6:L6").cellStyle.backColorRgb =
              Colors.amberAccent;
          sheet.getRangeByName("H1:L5").cellStyle.borders.all.lineStyle =
              LineStyle.thin;
          sheet.getRangeByName("H6:L6").cellStyle.borders.all.lineStyle =
              LineStyle.medium;
        }

        void title() {
          sheet.getRangeByName("H1").setText("Server Time (Minutes)");
          sheet.getRangeByName("I1").setText("Propability");
          sheet.getRangeByName("J1").setText("Cumulative");
          sheet.getRangeByName("L1").setText("To");
          sheet.getRangeByName("K1").setText("From");
          sheet.getRangeByName("H6:L6").merge();
          sheet.getRangeByName("H6:L6").setText("Server_1  Abdo");
        }

        void process() {
          for (int i = arrivalBetween[1]; i < arrivalBetween.length; i++) {
            sheet.getRangeByName("H${i}").setText('${arrivalBetween[i - 1]}');
          }

          for (int i = 0; i < ser_propability.length; i++) {
            sheet.getRangeByName("I${i + 2}").setNumber(ser_propability[i]);
          }

          for (int i = 0; i < ser_propability.length - 1; i++) {
            sheet.getRangeByName("J2").setFormula("${ser_propability[0]}");
            sheet
                .getRangeByName("J${i + 3}")
                .setFormula("=SUM(${"I${i + 3}"}+${"J${i + 2}"})");
          }
          for (int i = 0; i < ser_propability.length; i++) {
            sheet.getRangeByName("L${i + 2}").setFormula("=J${i + 2} * 100");
          }
          for (int i = 0; i < ser_propability.length - 1; i++) {
            sheet.getRangeByName("K2").setFormula("=0");
            sheet.getRangeByName("K${i + 3}").setFormula("=SUM(L${i + 2} + 1)");
          }
        }

        title();
        style();
        process();
      }

      void mostafaServer() {
        void style() {
          sheet.getRangeByName("N1:R1").cellStyle.backColorRgb = Colors.green;
          sheet.getRangeByName("N1:R1").cellStyle.borders.all.lineStyle =
              LineStyle.thin;
          sheet.getRangeByName("N1").columnWidth = 19;
          sheet.getRangeByName("O1").columnWidth = 12;
          sheet.getRangeByName("P1").columnWidth = 10;
          sheet.getRangeByName("N6:R6").cellStyle.backColorRgb =
              Colors.amberAccent;
          sheet.getRangeByName("N6:R6").cellStyle.borders.all.lineStyle =
              LineStyle.medium;
          sheet.getRangeByName("N1:R5").cellStyle.borders.all.lineStyle =
              LineStyle.thin;
          sheet.getRangeByName("A10:B10").cellStyle.hAlign = HAlignType.center;
          sheet.getRangeByName("A10:B10").cellStyle.borders.all.lineStyle =
              LineStyle.thick;
          sheet.getRangeByName("A10").cellStyle.backColorRgb =
              Colors.greenAccent;
          sheet.getRangeByName("B10").cellStyle.backColorRgb =
              Colors.orangeAccent;
        }

        void title() {
          sheet.getRangeByName("N1").setText("Server Time (Minutes)");
          sheet.getRangeByName("O1").setText("Propability");
          sheet.getRangeByName("P1").setText("Cumulative");
          sheet.getRangeByName("R1").setText("To");
          sheet.getRangeByName("Q1").setText("From");
          sheet.getRangeByName("N6:R6").merge();
          sheet.getRangeByName("N6:R6").setText("Server_2  Mostafa");
          sheet.getRangeByName("A10").setText("Opening");
          sheet.getRangeByName('B10').cellStyle.numberFormat = 'h:mm:ss AM/PM';
          sheet
              .getRangeByName('B10')
              .setDateTime(DateTime(2021, 8, 21, 8, 0, 0));
        }

        void process() {
          for (int i = arrivalBetween[2]; i <= arrivalBetween.length; i++) {
            sheet.getRangeByName("N${i - 1}").setFormula('${i}');
          }
          for (int i = 0; i < ser2_propability.length; i++) {
            sheet.getRangeByName("O${i + 2}").setNumber(ser2_propability[i]);
          }
          for (int i = 0; i < ser2_propability.length - 1; i++) {
            sheet.getRangeByName("P2").setFormula("${ser2_propability[0]}");
            sheet
                .getRangeByName("P${i + 3}")
                .setFormula("=SUM(${"O${i + 3}"}+${"P${i + 2}"})");
          }
          for (int i = 0; i < ser2_propability.length; i++) {
            sheet.getRangeByName("R${i + 2}").setFormula("=P${i + 2} * 100");
          }
          for (int i = 0; i < ser2_propability.length - 1; i++) {
            sheet.getRangeByName("Q2").setFormula("=0");
            sheet.getRangeByName("Q${i + 3}").setFormula("=SUM(R${i + 2} + 1)");
          }
        }

        title();
        style();
        process();
      }

      arrivalPropability();
      abdoServer();
      mostafaServer();
    }
    void sheet_2() {

        void style() {
          sheet2.getRangeByName("A2:E2").cellStyle.backColorRgb = Colors.amber;
          sheet2.getRangeByName("L2:M2").cellStyle.backColorRgb = Colors.amber;
          sheet2.getRangeByName("A1:M${number + 2}").cellStyle.hAlign =
              HAlignType.center;
          sheet2.getRangeByName("F1:H1").cellStyle.borders.all.lineStyle =
              LineStyle.thin;
          sheet2.getRangeByName("F1:H2").cellStyle.backColorRgb =
              Colors.teal.shade300;
          sheet2.getRangeByName("A3").columnWidth = 10;
          sheet2.getRangeByName("B2").columnWidth = 10.5;
          sheet2.getRangeByName("C2").columnWidth = 10.5;
          sheet2.getRangeByName("D2").columnWidth = 11;
          sheet2.getRangeByName("E2").columnWidth = 11;
          sheet2.getRangeByName("F2").columnWidth = 11;
          sheet2.getRangeByName("G2").columnWidth = 8;
          sheet2.getRangeByName("H2").columnWidth = 11;
          sheet2.getRangeByName("I1:K1").cellStyle.borders.all.lineStyle =
              LineStyle.thin;
          sheet2.getRangeByName("I1:K1").merge();
          sheet2.getRangeByName("I1:K2").cellStyle.backColorRgb =
              Colors.pinkAccent.shade200;
          sheet2.getRangeByName("I2").columnWidth = 11;
          sheet2.getRangeByName("M2").columnWidth = 10;
          sheet2.getRangeByName("J2").columnWidth = 8;
          sheet2.getRangeByName("K2").columnWidth = 11;



        }
        void title() {
          sheet2.getRangeByName("A2").setText("Customer ID");
          sheet2.getRangeByName("B2").setText("interval.rand");
          sheet2.getRangeByName("C2").setText("interval.time");
          sheet2.getRangeByName("D2").setText("Arrival.Clock");
          sheet2.getRangeByName("E2").setText("Service.rand");
          sheet2.getRangeByName("F1:H1").merge();
          sheet2.getRangeByName("F1:H1").setText("Abdo");

          sheet2.getRangeByName("F2").setText("Start");
          sheet2.getRangeByName("G2").setText("Duration");
          sheet2.getRangeByName("H2").setText("End");
          sheet2.getRangeByName("I1:K1").setText("Mostafa");
          sheet2.getRangeByName("I2").setText("Start");
          sheet2.getRangeByName("J2").setText("Duration");
          sheet2.getRangeByName("M2").setText("Server.Ideal");


        }
        void process() {

          for (int i = 1; i <= number; i++) {
            sheet2
                .getRangeByName("A${i + 1}:M${number + 2}")
                .cellStyle
                .borders
                .all
                .lineStyle = LineStyle.thin;
          }

          for (int i = 1; i <= number; i++) {
            sheet2.getRangeByName("A${i + 2}").setText('${i}');
          }
          for (int i = 1; i <= number; i++) {
            sheet2.getRangeByName("B${i + 2}").setFormula("=INT(RAND()*100 + 1)");
          }
          for (int i = 1; i <= number; i++) {
            sheet2.getRangeByName("C${i + 2}").setFormula(
                "=LOOKUP(B${i + 2},Sheet1!D${2}:E${5},Sheet1!A${2}:A${5})");
          }

          sheet2.getRangeByName("D3").setFormula("=Sheet1!B10+TIME(0,C3,0)");
          sheet2.getRangeByName('D3:D${number + 2}').cellStyle.numberFormat =
          'h:mm:ss AM/PM';
          for (int i = 1; i <= number - 1; i++) {
            sheet2
                .getRangeByName("D${i + 3}")
                .setFormula("=D${i + 2}+TIME(0,C${i + 3},0)");
          }

          for (int i = 1; i <= number; i++) {
            sheet2.getRangeByName("E${i + 2}").setFormula("=INT(RAND()*100 + 1)");
          }

          sheet2.getRangeByName('F3:F${number + 2}').cellStyle.numberFormat =
          'h:mm:ss AM/PM';
          sheet2.getRangeByName("F3").setFormula("D3");
          for (int i = 1; i < number; i++) {
            sheet2.getRangeByName("F${i + 3}").setFormula(
                "=IF(MAX(H${3}:H${i + 2})>MAX(K${3}:K${i + 2}),\'\',MAX(H${3}:H${i + 2},D${i + 3}))");
          }

          for (int i = 1; i <= number; i++) {
            sheet2.getRangeByName("G${i + 2}").setFormula(
                "=LOOKUP(E${i + 2},Sheet1!K${2}:L${5},Sheet1!H${2}:H${5})");
          }

          sheet2.getRangeByName("H3").setFormula("=F3+TIME(0,G3,0)");
          sheet2.getRangeByName('H3:H${number + 2}').cellStyle.numberFormat =
          'h:mm:ss AM/PM';
          for (int i = 1; i <= number; i++) {
            sheet2.getRangeByName("H${i + 3}").setFormula(
                "=IF(F${i + 3}<>\'\',F${i + 3}+TIME(0,G${i + 3},0),\'\')");
          }

          sheet2.getRangeByName('I4:I${number + 2}').cellStyle.numberFormat =
          'h:mm:ss AM/PM';
          for (int i = 1; i < number; i++) {
            sheet2.getRangeByName("I${i + 3}").setFormula(
                "=IF(F${i + 3}<>\'\',\'\',MAX(K${3}:K${i + 2},D${i + 3}))");
          }
          for (int i = 1; i <= number; i++) {
            sheet2.getRangeByName("J${i + 3}").setFormula(
                "=IF(I${i + 3}<>\'\',LOOKUP(E${i + 3},Sheet1!Q${2}:R${5},Sheet1!N${2}:N${5}),\'\')");
          }
          sheet2.getRangeByName("K2").setText("End");
          sheet2.getRangeByName('K4:K${number + 2}').cellStyle.numberFormat =
          'h:mm:ss AM/PM';
          for (int i = 1; i <= number; i++) {
            sheet2.getRangeByName("K${i + 3}").setFormula(
                "=IF(I${i + 3}<>\'\',I${i + 3}+TIME(0,J${i + 3},0),\'\')");
          }
          sheet2.getRangeByName("L2").setText("Customer.Waiting");
          sheet2.getRangeByName("L2").columnWidth = 15;
          for (int i = 1; i <= number; i++) {
            sheet2.getRangeByName("L${i + 2}").setFormula(
                "=MINUTE(IF(F${i + 2}<>\'\',F${i + 2}-D${i + 2},I${i + 2}-D${i + 2}))");
          }

          sheet2.getRangeByName("M2:M${number + 2}");
          for (int i = 1; i <= number; i++) {
            sheet2
                .getRangeByName("M${i + 2}")
                .setFormula("=IF(F${i + 2}<>\'\',\"Mostafa\",\"Abdo\")");
          }
        }
        title();
        process();
        style();

    }

    sheet_1();
    sheet_2();

    final List<int> bytes = workbook.saveAsStream();
    workbook.dispose();

    if (kIsWeb) {
      AnchorElement(
          href:
              'data:application/octet-stream;charset=utf-16le;base64,${base64.encode(bytes)}')
        ..setAttribute('download', 'Parallel_System.xlsx')
        ..click();
    } else {
      final String path = (await getApplicationSupportDirectory()).path;
      final String fileName = '$path/Parallel_System.xlsx';
      final File file = File(fileName);
      await file.writeAsBytes(bytes, flush: true);
      OpenFile.open(fileName);
    }
  }
}
