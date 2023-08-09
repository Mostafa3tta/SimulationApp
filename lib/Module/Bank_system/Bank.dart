import 'dart:io';
import 'dart:convert';
import 'package:flutter/material.dart';
import 'package:hexcolor/hexcolor.dart';
import 'package:open_file/open_file.dart';
import 'package:path_provider/path_provider.dart';
import 'package:syncfusion_flutter_xlsio/xlsio.dart';
import 'package:universal_html/html.dart' show AnchorElement;
import 'package:flutter/foundation.dart' show kIsWeb;

class Bank extends StatelessWidget {

  int number;

  Bank(this.number);

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

    // first table
    sheet.getRangeByName("A1:Q${number + 3}").cellStyle.hAlign = HAlignType.center;
    sheet.getRangeByName("A1:Q${number + 3}").cellStyle.vAlign = VAlignType.center;
    sheet.getRangeByName("B1").columnWidth = 8;
    sheet.getRangeByName("C1").columnWidth = 18.5;
    sheet.getRangeByName("D1").columnWidth = 16;
    sheet.getRangeByName("B2:B3").merge();
    sheet.getRangeByName("C2:C3").merge();
    sheet.getRangeByName("D2:D3").merge();
    sheet.getRangeByName("B2:B3").setText('Customer');
    sheet.getRangeByName("C2:C3").setText('Interarrival Times (a*)');
    sheet.getRangeByName("D2:D3").setText('Service Times (s*)');
    sheet.getRangeByName("B2:D3").cellStyle.borders.all.lineStyle =
        LineStyle.medium;
    sheet.getRangeByName("B2:D3").cellStyle.backColorRgb = Colors.orangeAccent.shade100;
    sheet.getRangeByName("B4:D${number + 3}").cellStyle.borders.all.lineStyle =
        LineStyle.thin;
    for (int i = 0; i < number; i++) {
      sheet.getRangeByName("B${i + 4}").setFormula('=${i+1}');
    }
    for (int i = 0; i < number; i++) {
      sheet.getRangeByName("C${i + 4}").setFormula("=INT(RAND()*11)"); //
    }
    for (int i = 0; i < number; i++) {
      sheet.getRangeByName("D${i + 4}").setFormula("=INT(RAND()*8)"); //
    }



    // simulation table
    sheet.getRangeByName("F2").columnWidth = 13.5;
    sheet.getRangeByName("G2").columnWidth = 10;
    sheet.getRangeByName("H2").columnWidth = 8;
    sheet.getRangeByName("I2").columnWidth = 12;
    sheet.getRangeByName("L3").columnWidth = 10;
    sheet.getRangeByName("M3").columnWidth = 12.5;
    sheet.getRangeByName("N3").columnWidth = 16;
    sheet.getRangeByName("O3").columnWidth = 16;
    sheet.getRangeByName("F2:F3").merge();
    sheet.getRangeByName("G2:G3").merge();
    sheet.getRangeByName("H2:H3").merge();
    sheet.getRangeByName("I2:I3").merge();
    sheet.getRangeByName("J2:K2").merge();
    sheet.getRangeByName("L2:M2").merge();
    sheet.getRangeByName("N2:O2").merge();
    sheet.getRangeByName("P2:Q2").merge();
    sheet.getRangeByName("F2:F3").setText('Interval/Service');
    sheet.getRangeByName("G2:G3").setText('Time Clock');
    sheet.getRangeByName("H2:H3").setText('Customer');
    sheet.getRangeByName("I2:I3").setText('Current Event');
    sheet.getRangeByName("J2:K2").setText('System State');
    sheet.getRangeByName("L2:M2").setText('Comments');
    sheet.getRangeByName("N2:O2").setText('FEL');
    sheet.getRangeByName("P2:Q2").setText('Cumulative Statistics');
    sheet.getRangeByName("J3").setText('LQ(t)');
    sheet.getRangeByName("K3").setText('LS(t)');
    sheet.getRangeByName("L3").setText('Next arrival');
    sheet.getRangeByName("M3").setText('Next daparture');
    sheet.getRangeByName("N3").setText('Order by value');
    sheet.getRangeByName("O3").setText('Order by logic');
    sheet.getRangeByName("P3").setText('B');
    sheet.getRangeByName("Q3").setText('MQ');
    sheet.getRangeByName("F2:Q3").cellStyle.borders.all.lineStyle =
        LineStyle.medium;
    sheet.getRangeByName("F2:Q3").cellStyle.backColorRgb = Colors.orangeAccent.shade100;
    sheet.getRangeByName("F4:Q${number + 3}").cellStyle.borders.all.lineStyle =
        LineStyle.thin;
    sheet.getRangeByName("F4").setFormula('=0');
    for (int i = 0; i < number - 1; i++) {
      sheet.getRangeByName("F${i + 5}").setFormula("=IF(I${i+5}='D',M${i+4},L${i+4})");
    }
    sheet.getRangeByName("G4").setFormula('=TIME(0,F4,0)+TIME(8,0,0)');
    sheet.getRangeByName('G4:G${number + 3}').cellStyle.numberFormat = 'h:mm AM/PM';
    for (int i = 0; i < number - 1; i++) {
      sheet.getRangeByName("G${i + 5}").setFormula("=TIME(0,F${i+5},0)+G${i+4}");
    }
    sheet.getRangeByName("H4").setFormula('=1');
    for (int i = 0; i < number - 1; i++) {
      sheet.getRangeByName("H${i + 5}").setFormula("=IF(I${i+5}='D',COUNTIF(I${4}:I${i+5},'D'),COUNTIF(I${4}:I${i+5},'A'))");
    }
    sheet.getRangeByName("I4").setText('A');
    for (int i = 0; i < number - 1; i++) {
      sheet.getRangeByName("I${i + 5}").setFormula("=IF(COUNTIF(I${4}:I${i+4},'A')=COUNTIF(I${4}:I${i+4},'D'),'A',IF(L${i+4}<M${i+4},'A','D'))");
    }
    sheet.getRangeByName("J4").setFormula('=0');
    for (int i = 0; i < number - 1; i++) {
      sheet.getRangeByName("J${i + 5}").setFormula("=COUNTIF(I${4}:I${i+5},'A')-COUNTIF(I${4}:I${i+5},'D')");
    }
    sheet.getRangeByName("K4").setFormula('=1');
    for (int i = 0; i < number - 1; i++) {
      sheet.getRangeByName("K${i + 5}").setFormula("=IF(COUNTIF(I${4}:I${i+5},'A')>COUNTIF(I${4}:I${i+5},'D'),1,0)");
    }
    sheet.getRangeByName("L4").setFormula('=C4');
    for (int i = 0; i < number - 1; i++) {
      sheet.getRangeByName("L${i + 5}").setFormula("=IF(I${i+5}='A',LOOKUP(H${i+5},B${4}:B${13},C${4}:C${13}),L${i+4})");
    }
    sheet.getRangeByName("M4").setFormula('=D4');
    for (int i = 0; i < number - 1; i++) {
      sheet.getRangeByName("M${i + 5}").setFormula("=IF(I${i+5}='D',LOOKUP(H${i+5}+1,B${4}:B${13},D${4}:D${13}),M${i+4})");
    }
    for (int i = 0; i < number; i++) {
      sheet.getRangeByName("N${i + 4}").setFormula("=IF(L${i+4}<=M${i+4}, CONCATENATE(\"(A,\",L${i+4},\"), (D,\",M${i+4},\"),(E,60)\"), CONCATENATE(\"(D,\",M${i+4},\"), (A,\",L${i+4},\"), (E,60)\"))");

    }
 for (int i = 0; i < number; i++) {
      sheet.getRangeByName("O${i + 4}").setFormula("=IF(COUNTIF(I\$${i+4}:I${i+4},\"A\")=COUNTIF(I\$${i+4}:I${i+4},\"D\"), CONCATENATE(\"(A,\",L${i+4},\"), (D,\",M${i+4},\"), (E,60)\"), IF(MIN(L${i+4}:M${i+4})=L${i+4},CONCATENATE(\"(A,\",L${i+4},\"), (D,\",M${i+4},\"), (E,60)\"), CONCATENATE(\"(D,\",M${i+4},\"), (A,\",L${i+4},\"), (E,60)\")))");

    }







    final List<int> bytes = workbook.saveAsStream();
    workbook.dispose();
    if (kIsWeb) {
      AnchorElement(
          href:
          'data:application/octet-stream;charset=utf-16le;base64},${base64.encode(bytes)}')
        ..setAttribute('download', 'Bank_System.xlsx')
        ..click();
    } else {
      final String path = (await getApplicationSupportDirectory()).path;
      final String fileName = '$path/Bank_System.xlsx';
      final File file = File(fileName);
      await file.writeAsBytes(bytes, flush: true);
      OpenFile.open(fileName);
    }
  }
}
