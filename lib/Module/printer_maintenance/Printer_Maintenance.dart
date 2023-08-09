import 'dart:convert';
import 'dart:io';
import 'dart:math';
import 'package:flutter/material.dart';
import 'package:hexcolor/hexcolor.dart';
import 'package:open_file/open_file.dart';
import 'package:path_provider/path_provider.dart';
import 'package:syncfusion_flutter_xlsio/xlsio.dart';
import 'package:universal_html/html.dart' show AnchorElement;
import 'package:flutter/foundation.dart' show kIsWeb;

class Printer_Maintenance extends StatelessWidget {
  int number;

  Printer_Maintenance(this.number);
  List DeviceLive = [1000, 1100, 1200, 1300, 1400, 1500, 1600, 1700];
  List DelayTime = [5, 10, 15];
  List<double> Probability = [];
  Random random = new Random();
  double sum = 0;

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

    // first table
    sheet.getRangeByName("B2:L11").cellStyle.hAlign = HAlignType.center;
    sheet.getRangeByName("B2:L11").cellStyle.vAlign = VAlignType.center;
    sheet.getRangeByName("B1").columnWidth = 16;
    sheet.getRangeByName("C1").columnWidth = 10;
    sheet.getRangeByName("D1").columnWidth = 10;
    sheet.getRangeByName("E1").columnWidth = 9;
    sheet.getRangeByName("F1").columnWidth = 9;
    sheet.getRangeByName("B2:B3").merge();
    sheet.getRangeByName("C2:C3").merge();
    sheet.getRangeByName("D2:D3").merge();
    sheet.getRangeByName("E2:F2").merge();
    sheet.getRangeByName("E2:F2").setText('Random.Digit');
    sheet.getRangeByName("E3").setText('From');
    sheet.getRangeByName("F3").setText('To');
    sheet.getRangeByName("B2:F3").cellStyle.borders.all.lineStyle =
        LineStyle.medium;
    sheet.getRangeByName("B4:F11").cellStyle.borders.all.lineStyle =
        LineStyle.thin;
    sheet.getRangeByName("B2:F3").cellStyle.backColorRgb = Colors.teal;
    sheet.getRangeByName("B2:B3").setText('Device Life (Hours)');
    sheet.getRangeByName("C2:C3").setText('Probability');
    sheet.getRangeByName("D2:D3").setText('Cumulative');
    for (int i = 0; i < DeviceLive.length; i++) {
      sheet.getRangeByName("B${i + 4}").setFormula('${DeviceLive[i]}');
    }
    for (int i = 0; i < DeviceLive.length; i++) {
      Probability.add(random.nextDouble());
      sum = sum + Probability[i];
    }
    for (int i = 0; i < Probability.length; i++) {
      Probability[i] = (Probability[i] * (1 / sum));
      sheet
          .getRangeByName("C${i + 4}")
          .setText("${Probability[i].toStringAsFixed(2)}");
    }
    sheet.getRangeByName("D4").setFormula("=C4");
    for (int i = 0; i < Probability.length - 1; i++) {
      sheet.getRangeByName("D${i + 5}").setFormula("=C${i + 5}+D${i + 4}");
    }
    sheet.getRangeByName("E4").setNumber(0);
    for (int i = 0; i < DeviceLive.length; i++) {
      sheet.getRangeByName("F${i + 4}").setFormula("=D${i + 4} * 100");
    }
    for (int i = 0; i < DeviceLive.length - 1 ; i++) {
      sheet.getRangeByName("E${i+5}").setFormula("=SUM(F${i+4}) + 1");
    }


    Probability = [];
    sum = 0;

    // second table
    sheet.getRangeByName("H1").columnWidth = 20;
    sheet.getRangeByName("I1").columnWidth = 10;
    sheet.getRangeByName("J1").columnWidth = 10;
    sheet.getRangeByName("K1").columnWidth = 9;
    sheet.getRangeByName("L1").columnWidth = 9;
    sheet.getRangeByName("H2:H3").merge();
    sheet.getRangeByName("I2:I3").merge();
    sheet.getRangeByName("J2:J3").merge();
    sheet.getRangeByName("K2:L2").merge();
    sheet.getRangeByName("K2:L2").setText('Random.Digit');
    sheet.getRangeByName("K3").setText('From');
    sheet.getRangeByName("L3").setText('To');
    sheet.getRangeByName("H2:L3").cellStyle.borders.all.lineStyle =
        LineStyle.medium;
    sheet.getRangeByName("H4:L6").cellStyle.borders.all.lineStyle =
        LineStyle.thin;
    sheet.getRangeByName("H2:L3").cellStyle.backColorRgb = Colors.amber;
    sheet.getRangeByName("H2:H3").setText('Delay Time (Minutes)');
    sheet.getRangeByName("I2:I3").setText('Probability');
    sheet.getRangeByName("J2:J3").setText('Cumulative');
    for (int i = 0; i < DelayTime.length; i++) {
      sheet.getRangeByName("H${i + 4}").setFormula('${DelayTime[i]}');
    }
    for (int i = 0; i < DelayTime.length; i++) {
      Probability.add(random.nextDouble());
      sum = sum + Probability[i];
    }
    for (int i = 0; i < Probability.length; i++) {
      Probability[i] = (Probability[i] * (1 / sum));
      sheet
          .getRangeByName("I${i + 4}")
          .setText("${Probability[i].toStringAsFixed(2)}");
    }
    sheet.getRangeByName("J4").setFormula("=I4");
    for (int i = 0; i < Probability.length - 1; i++) {
      sheet.getRangeByName("J${i + 5}").setFormula("=I${i + 5}+J${i + 4}");
    }
    sheet.getRangeByName("K4").setNumber(0);
    for (int i = 0; i < DelayTime.length; i++) {
      sheet.getRangeByName("L${i + 4}").setFormula("=J${i + 4} * 10");
    }
    for (int i = 0; i < DelayTime.length - 1 ; i++) {
      sheet.getRangeByName("K${i+5}").setFormula("=SUM(L${i+4}) + 1");
    }


    // simulation table
    sheet2.getRangeByName("A1:D1").merge();
    sheet2.getRangeByName("A3:D3").merge();
    sheet2.getRangeByName("A3:D3").setText('Rand Range for Delay');
    sheet2.getRangeByName("A2").setText('Max');
    sheet2.getRangeByName("A4").setText('Max');
    sheet2.getRangeByName("C2").setText('Min');
    sheet2.getRangeByName("C4").setText('Min');
    sheet2.getRangeByName("B2").setFormula('=100');
    sheet2.getRangeByName("D2").setFormula('=1');
    sheet2.getRangeByName("B4").setFormula('=10');
    sheet2.getRangeByName("D4").setFormula('=1');
    sheet2.getRangeByName("A1:D1").cellStyle.backColorRgb = Colors.yellow;
    sheet2.getRangeByName("A3:D3").cellStyle.backColorRgb = Colors.yellow;
    sheet2.getRangeByName("A1:D4").cellStyle.borders.all.lineStyle = LineStyle.thin;
    sheet2.getRangeByName("A2").cellStyle.backColorRgb = Colors.amber;
    sheet2.getRangeByName("A4").cellStyle.backColorRgb = Colors.amber;
    sheet2.getRangeByName("C2").cellStyle.backColorRgb = Colors.amber;
    sheet2.getRangeByName("C4").cellStyle.backColorRgb = Colors.amber;
    sheet2.getRangeByName("A1:D1").setText('Rand Range for Components');
    sheet2.getRangeByName("A1:L${number+13}").cellStyle.hAlign = HAlignType.center;
    sheet2.getRangeByName("A1:L${number+13}").cellStyle.vAlign = VAlignType.center;
    sheet2.getRangeByName("A6:L8").cellStyle.borders.all.lineStyle = LineStyle.medium;
    sheet2.getRangeByName("A9:L${number+8}").cellStyle.borders.all.lineStyle = LineStyle.thin;
    sheet2.getRangeByName("H6:L8").cellStyle.backColorRgb = Colors.pinkAccent;
    sheet2.getRangeByName("A6:A8").cellStyle.backColorRgb = Colors.pinkAccent;
    sheet2.getRangeByName("B6:C8").cellStyle.backColorRgb = Colors.green.shade300;
    sheet2.getRangeByName("D6:E8").cellStyle.backColorRgb = Colors.brown.shade300;
    sheet2.getRangeByName("F6:G8").cellStyle.backColorRgb = Colors.blue.shade300;
    sheet2.getRangeByName("A6:A8").merge();
    sheet2.getRangeByName("B6:C6").merge();
    sheet2.getRangeByName("B7:B8").merge();
    sheet2.getRangeByName("C7:C8").merge();
    sheet2.getRangeByName("D6:E6").merge();
    sheet2.getRangeByName("D7:D8").merge();
    sheet2.getRangeByName("E7:E8").merge();
    sheet2.getRangeByName("F6:G6").merge();
    sheet2.getRangeByName("F7:F8").merge();
    sheet2.getRangeByName("G7:G8").merge();
    sheet2.getRangeByName("H6:H8").merge();
    sheet2.getRangeByName("I6:I8").merge();
    sheet2.getRangeByName("J6:J8").merge();
    sheet2.getRangeByName("K6:K8").merge();
    sheet2.getRangeByName("L6:L8").merge();
    sheet2.getRangeByName("A6:A8").setText('ID');
    sheet2.getRangeByName("B6:C6").setText('Component 01');
    sheet2.getRangeByName("B7:B8").setText('R.D for Life');
    sheet2.getRangeByName("C7:C8").setText('Life (Hours)');
    sheet2.getRangeByName("D6:E6").setText('Component 02');
    sheet2.getRangeByName("D7:D8").setText('R.D for Life');
    sheet2.getRangeByName("E7:E8").setText('Life (Hours)');
    sheet2.getRangeByName("F6:G6").setText('Component 03');
    sheet2.getRangeByName("F7:F8").setText('R.D for Life');
    sheet2.getRangeByName("G7:G8").setText('Life (Hours)');
    sheet2.getRangeByName("H6:H8").setText('First Failure (Hours)');
    sheet2.getRangeByName("I6:I8").setText('Accumulated Life (Hours)');
    sheet2.getRangeByName("J6:J8").setText('R.D for Delay');
    sheet2.getRangeByName("K6:K8").setText('Delay (Minutes)');
    sheet2.getRangeByName("L6:L8").setText('Component');
    sheet2.getRangeByName("A6").columnWidth = 8;
    sheet2.getRangeByName("B6").columnWidth = 10;
    sheet2.getRangeByName("C6").columnWidth = 10;
    sheet2.getRangeByName("D6").columnWidth = 10;
    sheet2.getRangeByName("E6").columnWidth = 10;
    sheet2.getRangeByName("F6").columnWidth = 10;
    sheet2.getRangeByName("G6").columnWidth = 10;
    sheet2.getRangeByName("H6").columnWidth = 17;
    sheet2.getRangeByName("I6").columnWidth = 22;
    sheet2.getRangeByName("J6").columnWidth = 12;
    sheet2.getRangeByName("K6").columnWidth = 15;
    sheet2.getRangeByName("L6").columnWidth = 10;


    sheet2.getRangeByName('A9').setFormula('=1');
    for(int i = 0; i < number - 1; i++){
      sheet2.getRangeByName('A${i+10}').setFormula('=A${i+9}+1');
    }
    for(int i = 0; i < number; i++){
      sheet2.getRangeByName('B${i+9}').setFormula('=RANDBETWEEN(D${2},B${2})');
    }
    for(int i = 0; i < number; i++){
      sheet2.getRangeByName('C${i+9}').setFormula('=LOOKUP(Sheet2!B${i+9},Sheet1!E${4}:F${11},Sheet1!B${4}:B${11})');
    }
    for(int i = 0; i < number; i++){
      sheet2.getRangeByName('D${i+9}').setFormula('=RANDBETWEEN(D${2},B${2})');
    }
    for(int i = 0; i < number; i++){
      sheet2.getRangeByName('E${i+9}').setFormula('=LOOKUP(Sheet2!D${i+9},Sheet1!E${4}:F${11},Sheet1!B${4}:B${11})');
    }
    for(int i = 0; i < number; i++){
      sheet2.getRangeByName('F${i+9}').setFormula('=RANDBETWEEN(D${2},B${2})');
    }
    for(int i = 0; i < number; i++){
      sheet2.getRangeByName('G${i+9}').setFormula('=LOOKUP(Sheet2!F${i+9},Sheet1!E${4}:F${11},Sheet1!B${4}:B${11})');
    }
    for(int i = 0; i < number; i++){
      sheet2.getRangeByName('H${i+9}').setFormula('=MIN(C${i+9},E${i+9},G${i+9})');
    }
    sheet2.getRangeByName('I9').setFormula('=H9');
    for(int i = 0; i < number - 1; i++){
      sheet2.getRangeByName('I${i+10}').setFormula('=I${i+9}+H${i+10}');
    }
    for(int i = 0; i < number; i++){
      sheet2.getRangeByName('J${i+9}').setFormula('=RANDBETWEEN(D${4},B${4})');
    }
    for(int i = 0; i < number; i++){
      sheet2.getRangeByName('K${i+9}').setFormula('=LOOKUP(J${i+9},Sheet1!K${4}:L${6},Sheet1!H${4}:H${6})');
    }
    for(int i = 0; i < number; i++){
      sheet2.getRangeByName('L${i+9}').setFormula('=IF(H${i+9}=C${i+9},"Dev_01",IF(H${i+9}=E${i+9},"Dev_02","Dev_03"))');
    }

    sheet2.getRangeByName('J${number+9}').setText('Delay');
    sheet2.getRangeByName('J${number+9}').cellStyle.borders.all.lineStyle = LineStyle.medium;
    sheet2.getRangeByName('J${number+9}').cellStyle.backColorRgb = Colors.amber;
    sheet2.getRangeByName('K${number+9}').setFormula('=SUM(K9:K${number+8})');
    sheet2.getRangeByName('K${number+9}').cellStyle.borders.all.lineStyle = LineStyle.thin;
    sheet2.getRangeByName('K${number+9}').cellStyle.fontColorRgb = Colors.red;
    sheet2.getRangeByName('K${number+10}').setText('Dev_01');
    sheet2.getRangeByName('K${number+11}').setText('Dev_02');
    sheet2.getRangeByName('K${number+12}').setText('Dev_03');
    sheet2.getRangeByName('K${number+10}:K${number+12}').cellStyle.backColorRgb = Colors.amber;
    sheet2.getRangeByName('L${number+10}').setFormula('=COUNTIF(L${9}:L${number+8},"Dev_01")');
    sheet2.getRangeByName('L${number+11}').setFormula('=COUNTIF(L${9}:L${number+8},"Dev_02")');
    sheet2.getRangeByName('L${number+12}').setFormula('=COUNTIF(L${9}:L${number+8},"Dev_03")');
    sheet2.getRangeByName('K${number+10}:L${number+12}').cellStyle.borders.all.lineStyle = LineStyle.thin;




    Probability = [];
    sum = 0;




    final List<int> bytes = workbook.saveAsStream();
    workbook.dispose();

    if (kIsWeb) {
      AnchorElement(
          href:
          'data:application/octet-stream;charset=utf-16le;base64,${base64.encode(bytes)}')
        ..setAttribute('download', 'Printer_Maintenance_System.xlsx')
        ..click();
    } else {
      final String path = (await getApplicationSupportDirectory()).path;
      final String fileName = '$path/Printer_Maintenance_System.xlsx';
      final File file = File(fileName);
      await file.writeAsBytes(bytes, flush: true);
      OpenFile.open(fileName);
    }
  }
}
