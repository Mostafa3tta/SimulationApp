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

class Inventory_Screen extends StatelessWidget {
  late int number;

  Inventory_Screen(this.number);

  List Demand = [0, 1, 2, 3, 4];
  List Days = [1, 2, 3];
  Random random = new Random();
  List<double> Probability = [];
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
    sheet.getRangeByName("A1:L12").cellStyle.hAlign = HAlignType.center;
    sheet.getRangeByName("A1:L12").cellStyle.vAlign = VAlignType.center;
    sheet.getRangeByName("B1").columnWidth = 8;
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
    sheet.getRangeByName("B4:F8").cellStyle.borders.all.lineStyle =
        LineStyle.thin;
    sheet.getRangeByName("B2:F3").cellStyle.backColorRgb = Colors.yellowAccent;
    sheet.getRangeByName("B2:B3").setText('Demand');
    sheet.getRangeByName("C2:C3").setText('Probability');
    sheet.getRangeByName("D2:D3").setText('Cumulative');
    sheet.getRangeByName("D1:E1").merge();
    sheet.getRangeByName("D1:E1").setText("Demand");
    sheet.getRangeByName("D1:E1").cellStyle.borders.all.lineStyle =
        LineStyle.medium;
    sheet.getRangeByName("D1:E1").cellStyle.backColorRgb = Colors.red;
    for (int i = 0; i < Demand.length; i++) {
      sheet.getRangeByName("B${i + 4}").setText('${Demand[i]}');
    }
    for (int i = 0; i < Demand.length; i++) {
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
    for (int i = 0; i < Demand.length; i++) {
      sheet.getRangeByName("F${i + 4}").setFormula("=D${i + 4} * 100");
    }
    for (int i = 0; i < Demand.length - 1 ; i++) {
      sheet.getRangeByName("E${i+5}").setFormula("=SUM(F${i+4}) + 1");
    }

    Probability = [];   // very imprtant step
    sum = 0;

    // second table
    sheet.getRangeByName("H1").columnWidth = 14;
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
    sheet.getRangeByName("H2:L3").cellStyle.backColorRgb = Colors.yellowAccent;
    sheet.getRangeByName("H2:H3").setText('Lead Time (Days)');
    sheet.getRangeByName("I2:I3").setText('Probability');
    sheet.getRangeByName("J2:J3").setText('Cumulative');
    sheet.getRangeByName("I1:K1").merge();
    sheet.getRangeByName("I1:K1").setText("Lead Time");
    sheet.getRangeByName("I1:K1").cellStyle.borders.all.lineStyle =
        LineStyle.medium;
    sheet.getRangeByName("I1:K1").cellStyle.backColorRgb = Colors.red;
    for (int i = 0; i < Days.length; i++) {
      sheet.getRangeByName("H${i + 4}").setText('${Days[i]}');
    }
    for (int i = 0; i < Days.length; i++) {
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
    for (int i = 0; i < Days.length; i++) {
      sheet.getRangeByName("L${i + 4}").setFormula("=J${i + 4} * 100");
    }
    for (int i = 0; i < Days.length - 1 ; i++) {
      sheet.getRangeByName("K${i+5}").setFormula("=SUM(L${i+4}) + 1");
    }


    // third table
    sheet.getRangeByName("B11:C11").merge();
    sheet.getRangeByName("B11:C11").cellStyle.borders.all.lineStyle = LineStyle.medium;
    sheet.getRangeByName("B11:C11").cellStyle.hAlign = HAlignType.center;
    sheet.getRangeByName("B11:C11").cellStyle.backColorRgb = Colors.red;
    sheet.getRangeByName("B11:C11").setText('Limits');
    sheet.getRangeByName("B12:C12").cellStyle.borders.all.lineStyle = LineStyle.thin;
    sheet.getRangeByName("B12:C12").cellStyle.hAlign = HAlignType.center;
    sheet.getRangeByName("B12:C12").cellStyle.backColorRgb = Colors.yellowAccent;
    sheet.getRangeByName("B12").setText('Quantity Order');
    sheet.getRangeByName("B12").columnWidth = 13;
    sheet.getRangeByName("C12").setFormula('=10');



    // simulation table
    sheet2.getRangeByName("B2:K2").rowHeight = 35;
    sheet2.getRangeByName("B2:K2").cellStyle.hAlign = HAlignType.center;
    sheet2.getRangeByName("B2:K2").cellStyle.vAlign = VAlignType.center;
    sheet2.getRangeByName("B2:K2").cellStyle.borders.all.lineStyle = LineStyle.medium;
    sheet2.getRangeByName("B2:K2").cellStyle.backColorRgb = Colors.teal.shade300;
    sheet2.getRangeByName("B3:K3").cellStyle.backColorRgb = Colors.pinkAccent;
    sheet2.getRangeByName("B2").setText('Cycle');
    sheet2.getRangeByName("C2").setText('Day');
    sheet2.getRangeByName("D2").setText('Beginning Inventory');
    sheet2.getRangeByName("E2").setText('R.D for Demand');
    sheet2.getRangeByName("F2").setText('Demand');
    sheet2.getRangeByName("G2").setText('Ending Inventory');
    sheet2.getRangeByName("H2").setText('Shortage Quantity');
    sheet2.getRangeByName("I2").setText('Order Quantity');
    sheet2.getRangeByName("J2").setText('R.D for Lead Time');
    sheet2.getRangeByName("K2").setText('Days until Order arrives');
    sheet2.getRangeByName("B2").columnWidth = 8;
    sheet2.getRangeByName("C2").columnWidth = 8;
    sheet2.getRangeByName("D2").columnWidth = 18;
    sheet2.getRangeByName("E2").columnWidth = 14;
    sheet2.getRangeByName("F2").columnWidth = 8;
    sheet2.getRangeByName("G2").columnWidth = 15;
    sheet2.getRangeByName("H2").columnWidth = 17;
    sheet2.getRangeByName("I2").columnWidth = 14;
    sheet2.getRangeByName("J2").columnWidth = 15;
    sheet2.getRangeByName("K2").columnWidth = 20;
    sheet2.getRangeByName("B3:K${number + 2}").cellStyle.borders.all.lineStyle = LineStyle.thin;
    sheet2.getRangeByName("B3:K${number + 2}").cellStyle.hAlign = HAlignType.center;
    
    sheet2.getRangeByName('B4').setFormula('=1');
    sheet2.getRangeByName('C4').setFormula('=1');
    for(int i = 0; i < number - 2; i++){
      sheet2.getRangeByName('B${i+5}').setFormula('=IF(C${i+5}=1,MAX(B${4}:B${i+4})+1,"")');
    }
    for(int i = 0; i < number - 2; i++){
      sheet2.getRangeByName('C${i+5}').setFormula('=IF(C${i+4}>=5,1,C${i+4}+1)');
    }
    sheet2.getRangeByName('D3').setFormula('=1');
    for(int i = 0; i < number - 1; i++){
      sheet2.getRangeByName('D${i+4}').setFormula('=G${i+3}+IF(K${i+4}=0,10,0)');
    }
    for(int i = 0; i < number; i++){
      sheet2.getRangeByName('E${i+3}').setFormula('=INT(RAND()*100)');
    }
    for(int i = 0; i < number; i++){
      sheet2.getRangeByName('F${i+3}').setFormula('=LOOKUP(E${i+3},Sheet1!E${4}:F${8},Sheet1!B${4}:B${8})');
    }
    for(int i = 0; i < number; i++){
      sheet2.getRangeByName('G${i+3}').setFormula('=D${i+3}-F${i+3}');
    }
    sheet2.getRangeByName('H3').setFormula('=0');
    for(int i = 0; i < number - 1; i++){
      sheet2.getRangeByName('H${i+4}').setFormula('=IF((H${i+3}+F${i+4}-D${i+4})>0,(H${i+3}+F${i+4}-D${i+4}),0)');
    }
    for(int i = 0; i < number - 2; i++){
      sheet2.getRangeByName('I${i+4}').setFormula('=IF(C${i+4}=1,10,"")');
    }
    for(int i = 0; i < number; i++){
      sheet2.getRangeByName('J${i+3}').setFormula('=IF(I${i+3}<>"",INT((RAND()*100)),"")');
    }
    for(int i = 0; i < number - 2; i++){
      sheet2.getRangeByName('K${i+4}').setFormula('=IF(J${i+4}<>"",LOOKUP(J${i+4},Sheet1!K${4}:L${6},Sheet1!H${4}:H${6}'
          '),IF(K${i+3}<>"",IF(K${i+3}>0,K${i+3}-1,""),""))');
    }

    Probability = [];   // very imprtant step
    sum = 0;

    final List<int> bytes = workbook.saveAsStream();
    workbook.dispose();

    if (kIsWeb) {
      AnchorElement(
          href:
              'data:application/octet-stream;charset=utf-16le;base64,${base64.encode(bytes)}')
        ..setAttribute('download', 'Inventory_System.xlsx')
        ..click();
    } else {
      final String path = (await getApplicationSupportDirectory()).path;
      final String fileName = '$path/Inventory_System.xlsx';
      final File file = File(fileName);
      await file.writeAsBytes(bytes, flush: true);
      OpenFile.open(fileName);
    }
  }
}
