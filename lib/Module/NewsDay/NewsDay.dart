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

class NewsDay_Screen extends StatelessWidget {
  late int number;

  NewsDay_Screen(this.number);

  List<String> TON = [
    'Good',
    'Fair',
    'poor',
  ];
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


    void sheet1(){
      sheet.getRangeByName("A1:Q11").cellStyle.hAlign = HAlignType.center;
      sheet.getRangeByName("A1:Q11").cellStyle.vAlign = VAlignType.center;




      void table1() {

        void style(){
          sheet.getRangeByName("B2").columnWidth = 17;
          sheet.getRangeByName("C2").columnWidth = 10;
          sheet.getRangeByName("D2").columnWidth = 10;
          sheet.getRangeByName("E2").columnWidth = 10;
          sheet.getRangeByName("F2").columnWidth = 10;
          sheet.getRangeByName("D1:E1").cellStyle.backColorRgb = Colors.redAccent;
          sheet.getRangeByName("D1:E1").cellStyle.borders.all.lineStyle =
              LineStyle.medium;
          sheet.getRangeByName("B2:F3").cellStyle.backColorRgb =
              Colors.yellowAccent;
          sheet.getRangeByName("B2:F6").cellStyle.borders.all.lineStyle=LineStyle.thin;
        }

        style();
        //First Table

        sheet.getRangeByName("D1:E1").merge();
        sheet.getRangeByName("D1:E1").setText("NewsDay Type");


        sheet.getRangeByName("B2:B3").merge();
        sheet.getRangeByName("B2:B3").setText("Type of Newsdays");

        for (int i = 0; i < TON.length; i++) {
          sheet.getRangeByName("B${i + 4}").setText('${TON[i]}');
        }

        sheet.getRangeByName("C2:C3").merge();
        sheet.getRangeByName("C2:C3").setText("Probability");

        for (int i = 0; i < 3; i++) {
          Probability.add(random.nextDouble());
          sum = sum + Probability[i];
        }
        for (int i = 0; i < Probability.length; i++) {
          Probability[i] = (Probability[i] * (1 / sum));
          sheet
              .getRangeByName("C${i + 4}")
              .setText("${Probability[i].toStringAsFixed(2)}");
        }

        sheet.getRangeByName("D2:D3").merge();
        sheet.getRangeByName("D2:D3").setText("Cumulative");

        sheet.getRangeByName("D4").setFormula("=C4");
        for (int i = 0; i < Probability.length - 1; i++) {
          sheet.getRangeByName("D${i + 5}").setFormula("=C${i + 5}+D${i + 4}");
        }

        sheet.getRangeByName("E2:E3").merge();
        sheet.getRangeByName("E2:E3").setText("From");
        sheet.getRangeByName("B2:F3").cellStyle.borders.all.lineStyle =
            LineStyle.thin;
        sheet.getRangeByName("E4").setNumber(0);

        for (int i = 0; i < TON.length; i++) {
          sheet.getRangeByName("F${i + 4}").setFormula("=D${i + 4} * 100");
        }

        sheet.getRangeByName("F2:F3").merge();
        sheet.getRangeByName("F2:F3").setText("To");

        for (int i = 0; i < TON.length - 1; i++) {
          sheet.getRangeByName("E${i + 5}").setFormula("=SUM(F${i + 4}) + 1");
        }
      }

      table1();

      Probability = [];
      sum = 0;

      void table2() {
        //Style ٍSecond Table
        void style(){
          sheet.getRangeByName("B11:C14").cellStyle.backColorRgb =
              Colors.yellowAccent;

          sheet.getRangeByName("B10:C10").merge();
          sheet.getRangeByName("B10:C10").setText("Limits");


          sheet.getRangeByName("B10:C14").cellStyle.borders.all.lineStyle=LineStyle.medium;
          sheet.getRangeByName("B10:C14").cellStyle.hAlign = HAlignType.center;
          sheet.getRangeByName("B10:C14").cellStyle.vAlign = VAlignType.center;
        }
        style();
        //Second Table
        sheet.getRangeByName("B11").setText("Days");
        sheet.getRangeByName("C11").setNumber(20);
        sheet.getRangeByName("B12").setText("Quantity");
        sheet.getRangeByName("C12").setNumber(70);
        sheet.getRangeByName("B13").setText("Cost");
        sheet.getRangeByName("C13").setNumber(0.17);
        sheet.getRangeByName("B14").setText("Salvage");
        sheet.getRangeByName("C14").setNumber(0.05);
      }

      table2();



      void table3(){
        sheet.getRangeByName("K1:M1").merge();
        sheet.getRangeByName("K1:M1").setText("NewsDay Distribution");
        sheet.getRangeByName("K1:M1").cellStyle.backColorRgb = Colors.redAccent;
        sheet.getRangeByName("K1:M1").cellStyle.borders.all.lineStyle=LineStyle.medium;
        sheet.getRangeByName("H2:Q11").cellStyle.borders.all.lineStyle=LineStyle.thin;

        List <double> good = [
          0.03,
          0.08,
          0.23,
          0.43,
          0.78,
          0.93,
          1,
        ];
        List <double> fair = [
          0.1,
          0.28,
          0.68,
          0.88,
          0.96,
          1,
          1,
        ];
        List <double> poor = [
          0.44,
          0.66,
          0.82,
          0.94,
          1,
          1,
          1
        ];

        void style(){
          sheet.getRangeByName("B10:C10").cellStyle.backColorRgb = Colors.redAccent;
          sheet.getRangeByName("B10:C10").cellStyle.borders.all.lineStyle =
              LineStyle.medium;

          sheet.getRangeByName("H2:H4").merge();
          sheet.getRangeByName("H2:H4").setText("Demand");
          sheet.getRangeByName("H2:H4").cellStyle.backColorRgb = Colors.yellowAccent;


          sheet.getRangeByName("I2:K2").merge();
          sheet.getRangeByName("I2:K2").setText("Cumulative Distribution");
          sheet.getRangeByName("I2:K2").cellStyle.backColorRgb = Colors.yellowAccent;


          sheet.getRangeByName("I3:I4").merge();
          sheet.getRangeByName("I3:I4").setText("Good");
          sheet.getRangeByName("I3:I11").cellStyle.backColorRgb = Colors.blue.shade300;
          sheet.getRangeByName("L3:M11").cellStyle.backColorRgb = Colors.blue.shade300;
          sheet.getRangeByName("I3:I4").cellStyle.borders.all;

          sheet.getRangeByName("J3:J4").merge();
          sheet.getRangeByName("J3:J4").setText("Fair");
          sheet.getRangeByName("J3:J11").cellStyle.backColorRgb = Colors.purple.shade300;
          sheet.getRangeByName("N3:O11").cellStyle.backColorRgb = Colors.purple.shade300;
          sheet.getRangeByName("I3:I4").cellStyle.borders.all;



          sheet.getRangeByName("K3:K4").merge();
          sheet.getRangeByName("K3:K4").setText("Poor");
          sheet.getRangeByName("K3:K11").cellStyle.backColorRgb = Colors.deepOrange.shade300;
          sheet.getRangeByName("P3:Q11").cellStyle.backColorRgb = Colors.deepOrange.shade300;
          sheet.getRangeByName("K3:K4").cellStyle.borders.all;


          sheet.getRangeByName("L2:Q2").merge();
          sheet.getRangeByName("L2:Q2").setText("Random.Digit");
          sheet.getRangeByName("L2:Q2").cellStyle.backColorRgb = Colors.yellowAccent;
          sheet.getRangeByName("L2:Q2").cellStyle.borders.all;




        }
        style();
//Third Table


        double number =40;
        for (int i=5;i<=11;i++)
        {
          sheet.getRangeByName("H$i").setNumber(number);
          number+=10;

        }



        void cumulative (){


          for (int i=0;i<good.length;i++)
          {
            sheet.getRangeByName("i${i+5}").setNumber(good[i]);
          }

          for (int i=0;i<fair.length;i++)
          {
            sheet.getRangeByName("j${i+5}").setNumber(fair[i]);
          }

          for (int i=0;i<poor.length;i++)
          {
            sheet.getRangeByName("k${i+5}").setNumber(poor[i]);
          }
        }

        cumulative();



        void random (){
          void goodCell (){
            sheet.getRangeByName("L3:M3").merge();
            sheet.getRangeByName("L3:M3").setText("Good ");


            sheet.getRangeByName("M4").setText("To ");

            for (int i = 0; i < good.length ; i++) {
              sheet.getRangeByName("M${i +5}").setFormula("=I${i+5} * 100");

            }


            sheet.getRangeByName("L4").setText("From ");

            sheet.getRangeByName("L5").setNumber(0);
            for (int i = 0; i < good.length-1; i++) {
              sheet.getRangeByName("L${i + 6}").setFormula("=SUM(M${i + 5}) + 1");
            }


          }
          goodCell();

          void fairCell (){
            sheet.getRangeByName("N3:O3").merge();
            sheet.getRangeByName("N3:O3").setText("Fair ");


            sheet.getRangeByName("O4").setText("To ");

            for (int i = 0; i < fair.length ; i++) {
              sheet.getRangeByName("O${i +5}").setFormula("=J${i+5} * 100");

            }


            sheet.getRangeByName("N4").setText("From ");

            sheet.getRangeByName("N5").setNumber(0);
            for (int i = 0; i < fair.length-1; i++) {
              sheet.getRangeByName("N${i + 6}").setFormula("=SUM(O${i + 5}) + 1");
            }


          }
          fairCell();

          void poorCell (){
            sheet.getRangeByName("P3:Q3").merge();
            sheet.getRangeByName("P3:Q3").setText("Poor ");


            sheet.getRangeByName("Q4").setText("To ");

            for (int i = 0; i < good.length ; i++) {
              sheet.getRangeByName("Q${i +5}").setFormula("=K${i+5} * 100");

            }


            sheet.getRangeByName("P4").setText("From ");

            sheet.getRangeByName("P5").setNumber(0);
            for (int i = 0; i < good.length-1; i++) {
              sheet.getRangeByName("P${i + 6}").setFormula("=SUM(Q${i + 5}) + 1");
            }


          }
          poorCell();

        }
        random();


      }
      table3();

    }
    sheet1();



    void sheetSimulation(){
      sheet2.getRangeByName("B3:N${number+4}").cellStyle.hAlign = HAlignType.center;
      sheet2.getRangeByName("B3:N${number+4}").cellStyle.vAlign = VAlignType.center;
      sheet2.getRangeByName("B3:N${number+4}").cellStyle.borders.all.lineStyle=LineStyle.thin;

      void title (){

        sheet2.getRangeByName("B3:N4").cellStyle.hAlign = HAlignType.center;
        sheet2.getRangeByName("B3:N4").cellStyle.vAlign = VAlignType.center;
        sheet2.getRangeByName("B3:N4").cellStyle.borders.all;
        sheet2.getRangeByName("B3:N4").cellStyle.backColorRgb=Colors.orange;

        sheet2.getRangeByName("B3").columnWidth = 6;
        sheet2.getRangeByName("C3").columnWidth = 22;
        sheet2.getRangeByName("D3").columnWidth = 15;
        sheet2.getRangeByName("E3").columnWidth = 15;
        sheet2.getRangeByName("F3").columnWidth = 8;
        sheet2.getRangeByName("G3").columnWidth = 18;
        sheet2.getRangeByName("H3").columnWidth = 28;
        sheet2.getRangeByName("I3").columnWidth = 21;
        sheet2.getRangeByName("J3").columnWidth = 10;
        sheet2.getRangeByName("K3").columnWidth = 20;
        sheet2.getRangeByName("L3").columnWidth = 9;
        sheet2.getRangeByName("M3").columnWidth = 8;
        sheet2.getRangeByName("N3").columnWidth = 11;



        sheet2.getRangeByName("B3:B4").merge();
        sheet2.getRangeByName("B3:B4").setText("Day");

        sheet2.getRangeByName("C3:C4").merge();
        sheet2.getRangeByName("C3:C4").setText("R.D.for Type of NewsDay");


        sheet2.getRangeByName("D3:D4").merge();
        sheet2.getRangeByName("D3:D4").setText("Type of NewsDay");

        sheet2.getRangeByName("E3:E4").merge();
        sheet2.getRangeByName("E3:E4").setText("R.D.for Demand");


        sheet2.getRangeByName("F3:F4").merge();
        sheet2.getRangeByName("F3:F4").setText("Demand");


        sheet2.getRangeByName("G3:G4").merge();
        sheet2.getRangeByName("G3:G4").setText("Revenue From Sales");


        sheet2.getRangeByName("H3:H4").merge();
        sheet2.getRangeByName("H3:H4").setText("Lost Profit From Excess Demand");


        sheet2.getRangeByName("I3:I4").merge();
        sheet2.getRangeByName("I3:I4").setText("Salvage From Scrap Sale");

        sheet2.getRangeByName("J3:J4").merge();
        sheet2.getRangeByName("J3:J4").setText("Daily Profit");

        sheet2.getRangeByName("K3:K4").merge();
        sheet2.getRangeByName("K3:K4").setText("remaining newspapers");

        sheet2.getRangeByName("L3:L4").merge();
        sheet2.getRangeByName("L3:L4").setText("Net Profit");

        sheet2.getRangeByName("M3:M4").merge();
        sheet2.getRangeByName("M3:M4").setText("True Lost");

        sheet2.getRangeByName("N3:N4").merge();
        sheet2.getRangeByName("N3:N4").setText("Full True Lost");

      }
      title ();
//Day
         void day() {
        for (int i = 1; i <= number; i++) {
          sheet2.getRangeByName("B${i + 4}").setText(i.toString());
        }
      }
//R.D
        void randForType (){
        for (int i = 1; i <= number; i++) {
          sheet2.getRangeByName("C${i + 4}").setFormula("=INT(RAND()*100)");
        }
      }
//Types of News Days
          void typesOfNewsDays(){
        for (int i = 1; i <= number; i++) {
          sheet2.getRangeByName("D${i + 4}").setFormula(
              "=LOOKUP(C${i + 4},Sheet1!E${4}:F${6},Sheet1!B${4}:B${6})");
        }
      }
//R.D.for Demand
      void randForDemand() {
        for (int i = 1; i <= number; i++) {
          sheet2
              .getRangeByName("E${i + 4}")
              .setFormula("==ROUND(RAND()*100,0)");
        }
      }
//Demand
      void demand(){
        for (int i = 1; i <= number; i++) {
          sheet2.getRangeByName("F${i + 4}").setFormula(
          "=IF(D${i+4}=\"Good\",LOOKUP(E${i+4},Sheet1!L${5}:M${11},Sheet1!H${5}:H${11}),IF(D${i+4}=\"Fair\",LOOKUP(E${i+4},Sheet1!N${5}:O${11},Sheet1!H${5}:H${11}),LOOKUP(E${i+4},Sheet1!P${5}:Q${11},Sheet1!H${5}:H${11})))"
          );
        }
      }
//Revenue From Sales
      void revenueFromSales() {
        for (int i = 1; i <= number; i++) {
          sheet2.getRangeByName("G${i + 4}").setFormula(
              "=IF(F${i + 4}<Sheet1!C${12},F${i +
                  4}*Sheet1!C${13},Sheet1!C${12}*Sheet1!C${13})"
          );
        }
      }
//Lost Profit From Excess Demand
      void lostProfitFromExcessDemand(){
        for (int i = 1; i <= number; i++) {
          sheet2.getRangeByName("H${i + 4}").setFormula(
              "=IF(F${i + 4}>Sheet1!C${12},(F${i +
                  4}-Sheet1!C${12})*Sheet1!C${13},0)"
          );
        }
      }

//Salvage From Scrap Sale
      void salvageFromScrapSale(){
        for (int i = 1; i <= number; i++) {
          sheet2.getRangeByName("I${i + 4}").setFormula(
            "=IF(F${i+4}<Sheet1!C${12},(Sheet1!C${12}-F${i+4})*Sheet1!C${14},0)"
          );
        }
      }
//Daily Profit
      void dailyProfit(){
        for (int i = 1; i <= number; i++) {
          sheet2.getRangeByName("J${i + 4}").setFormula(
            "=G${i+4}+I${i+4}-H${i+4}"
          );
        }
      }
//remaining newspapers
      void RemainingNewsPapers(){

        for (int i = 1; i <= number; i++) {
          sheet2.getRangeByName("K${i + 4}").setFormula(
              "=IF(Sheet1!C${12}>F${i+4},Sheet1!C${12}-F${i+4},0)"
          );
        }

      }
//Net Profit
      void netProfit(){
        for (int i = 1; i <= number; i++) {
          sheet2.getRangeByName("L${i + 4}").setFormula(
            "=G${i+4}+I${i+4}"
          );
        }

      }
//True Lost
      void trueLost(){

        for (int i = 1; i <= number; i++) {
          sheet2.getRangeByName("M${i + 4}").setFormula(
            "=IF(F${i+4}<Sheet1!C${12},(Sheet1!C${12}-F${i+4})*(Sheet1!C${13}-Sheet1!C${14}),0)"
          );
        }
      }
//Full True Lost

      void fullTrueLost(){
        for (int i = 1; i <= number; i++) {
          sheet2.getRangeByName("N${i + 4}").setFormula(
            "=M${i+4}+H${i+4}"
          );
        }
      }

      day();
      randForType();
      typesOfNewsDays();
      randForDemand();
      demand();
      revenueFromSales();
      lostProfitFromExcessDemand();
      salvageFromScrapSale();
      dailyProfit();
      RemainingNewsPapers();
      netProfit();
      trueLost();
      fullTrueLost();

    }
    sheetSimulation();


    Probability = [];
    sum = 0;


    final List<int> bytes = workbook.saveAsStream();
    workbook.dispose();

    if (kIsWeb) {
      AnchorElement(
          href:
              'data:application/octet-stream;charset=utf-16le;base64,${base64.encode(bytes)}')
        ..setAttribute('download', 'NewsDay_System.xlsx')
        ..click();
    } else {
      final String path = (await getApplicationSupportDirectory()).path;
      final String fileName = '$path/NewsDay_System.xlsx';
      final File file = File(fileName);
      await file.writeAsBytes(bytes, flush: true);
      OpenFile.open(fileName);
    }
  }
}
