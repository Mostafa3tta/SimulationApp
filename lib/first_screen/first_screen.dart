import 'package:animated_text_kit/animated_text_kit.dart';
import 'package:flutter/cupertino.dart';
import 'package:flutter/material.dart';
import 'package:flutter_animated_button/flutter_animated_button.dart';
import 'package:simulation_excel/Module/Bank_system/Bank.dart';
import 'package:simulation_excel/Module/NewsDay/NewsDay.dart';
import 'package:simulation_excel/Module/Perial_sirver/perial_server.dart';
import 'package:simulation_excel/Module/invintory/inventory.dart';
import 'package:simulation_excel/Module/printer_maintenance/Printer_Maintenance.dart';
import 'package:simulation_excel/components/components.dart';
import '../Module/single_sirver/single_server.dart';
import 'package:hexcolor/hexcolor.dart';
import 'package:marquee/marquee.dart';

class FirstScreen extends StatefulWidget {
  @override
  _FirstScreenState createState() => _FirstScreenState();
}

class _FirstScreenState extends State<FirstScreen> {
  var customerNumberController = TextEditingController();

  var arrivalStartController = TextEditingController();

  var endTimeController = TextEditingController();

  int num = 0;
  var formKey = GlobalKey<FormState>();

  var isPassword = true;

  Widget buildContainer(Widget child) {
    return Container(
      decoration: BoxDecoration(
        color: Colors.teal[100],
        border: Border.all(color: Colors.blue),
        borderRadius: BorderRadius.circular(10),
      ),
      margin: EdgeInsets.all(10),
      padding: EdgeInsets.all(10),
      height: 200,
      width: 350,
      child: child,
    );
  }

  void show() {
      ScaffoldMessenger.of(context).showSnackBar(SnackBar(
        behavior: SnackBarBehavior.floating,
        content: Row(
          mainAxisAlignment: MainAxisAlignment.spaceBetween,
          children: [
            Text(
              'You must Enter Customer Number',
              style: TextStyle(
                color: Colors.black,
                fontWeight: FontWeight.bold,
              ),
            ),
            Icon(Icons.format_list_numbered_rounded),
          ],
        ),
        backgroundColor: Colors.blue.shade400,
        duration: Duration(seconds: 3),
      ));
  }

  @override
  Widget build(BuildContext context) {
    List buildButton = [
      AnimatedButton(
          height: 50,
          width: 250,
          text: "Create Single Server",
          isReverse: true,
          selectedTextColor: Colors.black,
          transitionType: TransitionType.LEFT_TO_RIGHT,
          textStyle: const TextStyle(
            fontSize: 20,
          ),
          backgroundColor: HexColor("#6fa8dc"),
          selectedBackgroundColor: Colors.lightBlueAccent,
          borderRadius: 50,
          borderWidth: 5,
          onPress: () async {
            if(customerNumberController.text.isEmpty){
              show();
            }
            num = int.parse(customerNumberController.text);
            navigateTo(context, homePageScreen(num));
          }),
      AnimatedButton(
          height: 50,
          width: 250,
          text: "Create Parallel Server",
          isReverse: true,
          selectedTextColor: Colors.black,
          transitionType: TransitionType.LEFT_TO_RIGHT,
          textStyle: const TextStyle(
            fontSize: 20,
          ),
          backgroundColor: HexColor("#6fa8dc"),
          selectedBackgroundColor: Colors.lightBlueAccent,
          borderRadius: 50,
          borderWidth: 5,
          onPress: () {
            if(customerNumberController.text.isEmpty){
              show();
            }
            num = int.parse(customerNumberController.text);
            navigateTo(context, Perial_Screen(num));
          }),
      AnimatedButton(
          height: 50,
          width: 250,
          text: "Create Inventory Server",
          isReverse: true,
          selectedTextColor: Colors.black,
          transitionType: TransitionType.LEFT_TO_RIGHT,
          textStyle: const TextStyle(
            fontSize: 20,
          ),
          backgroundColor: HexColor("#6fa8dc"),
          selectedBackgroundColor: Colors.lightBlueAccent,
          borderRadius: 50,
          borderWidth: 5,
          onPress: () {
            if(customerNumberController.text.isEmpty){
              show();
            }
            num = int.parse(customerNumberController.text);
            navigateTo(context, Inventory_Screen(num));
          }),
      AnimatedButton(
          height: 50,
          width: 250,
          text: "Create NewsDay Server",
          isReverse: true,
          selectedTextColor: Colors.black,
          transitionType: TransitionType.LEFT_TO_RIGHT,
          textStyle: const TextStyle(
            fontSize: 20,
          ),
          backgroundColor: HexColor("#6fa8dc"),
          selectedBackgroundColor: Colors.lightBlueAccent,
          borderRadius: 50,
          borderWidth: 5,
          onPress: () {
            if(customerNumberController.text.isEmpty){
              show();
            }
            num = int.parse(customerNumberController.text);
            navigateTo(context, NewsDay_Screen(num));
          }),
      AnimatedButton(
          height: 50,
          width: 250,
          text: "Printer Maintenance",
          isReverse: true,
          selectedTextColor: Colors.black,
          transitionType: TransitionType.LEFT_TO_RIGHT,
          textStyle: const TextStyle(
            fontSize: 20,
          ),
          backgroundColor: HexColor("#6fa8dc"),
          selectedBackgroundColor: Colors.lightBlueAccent,
          borderRadius: 50,
          borderWidth: 5,
          onPress: () {
            if(customerNumberController.text.isEmpty){
              show();
            }
            num = int.parse(customerNumberController.text);
            navigateTo(context, Printer_Maintenance(num));
          }),
      AnimatedButton(
          height: 50,
          width: 250,
          text: "Create Bank System",
          isReverse: true,
          selectedTextColor: Colors.black,
          transitionType: TransitionType.LEFT_TO_RIGHT,
          textStyle: const TextStyle(
            fontSize: 20,
          ),
          backgroundColor: HexColor("#6fa8dc"),
          selectedBackgroundColor: Colors.lightBlueAccent,
          borderRadius: 50,
          borderWidth: 5,
          onPress: () {
            if(customerNumberController.text.isEmpty){
              show();
            }
            num = int.parse(customerNumberController.text);
            navigateTo(context, Bank(num));
          }),
    ];

    return Container(
      decoration: BoxDecoration(
          gradient: LinearGradient(
              begin: Alignment.topLeft,
              end: Alignment.bottomRight,
              colors: [
            HexColor("#acb6e5"),
            HexColor("#86fde8"),
          ])),
      child: Scaffold(
        backgroundColor: Colors.transparent,
        body: Padding(
          padding: const EdgeInsets.all(15.0),
          child: SingleChildScrollView(
            child: Form(
              key: formKey,
              child: Column(
                mainAxisSize: MainAxisSize.min,
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  const SizedBox(height: 50),
                  const Text(
                    'Be',
                    style: TextStyle(
                      fontSize: 30.0,
                    ),
                  ),
                  const SizedBox(width: 20.0, height: 30.0),
                  Container(
                    width: double.infinity,
                    height: 20,
                    child: AnimatedTextKit(repeatForever: true, animatedTexts: [
                      RotateAnimatedText('Simulation',
                          alignment: Alignment.center,
                          textStyle: const TextStyle(
                            fontSize: 15,
                            color: Colors.red,
                            fontWeight: FontWeight.bold,
                          )),
                      RotateAnimatedText('Your',
                          alignment: Alignment.center,
                          textStyle: const TextStyle(
                            color: Colors.green,
                            fontSize: 15,
                            fontWeight: FontWeight.bold,
                          )),
                      RotateAnimatedText('Life',
                          alignment: Alignment.center,
                          textStyle: const TextStyle(
                            color: Colors.yellowAccent,
                            fontSize: 15,
                            fontWeight: FontWeight.bold,
                          )),
                    ]),
                  ),
                  Text(
                    "Let's Do our Business Easily",
                    style: Theme.of(context).textTheme.subtitle1!.copyWith(
                          fontSize: 15.0,
                        ),
                    maxLines: 1,
                  ),
                  const SizedBox(height: 40),
                  Padding(
                    padding: const EdgeInsets.all(8.0),
                    child: Text(
                      "Number of Customer ",
                      style: Theme.of(context).textTheme.subtitle1!.copyWith(
                            fontSize: 15.0,
                            fontWeight: FontWeight.bold,
                          ),
                      maxLines: 1,
                    ),
                  ),
                  Center(
                    child: defaultFormField(
                      onTap: () {},
                      controller: customerNumberController,
                      type: TextInputType.number,
                      onChange: () {},
                      isVaildator: true,
                      validate: () {},
                      label: "Customer Number",
                      prefix: Icons.format_list_numbered,
                      hint: "Enter the Number here",
                    ),
                  ),
                  const SizedBox(height: 40),
                  Container(
                    alignment: Alignment.center,
                    child: Text(
                      "Choose the server you want to create from here:",
                      style: Theme.of(context).textTheme.subtitle1!.copyWith(
                            fontSize: 15,
                            fontWeight: FontWeight.bold,
                          ),
                      maxLines: 1,
                    ),
                  ),
                  //const SizedBox(height: 5),
                  Container(
                    alignment: Alignment.center,
                    child: buildContainer(
                      ListView.builder(
                        itemBuilder: (ctx, index) => Column(
                          children: [
                            ListTile(
                              leading: CircleAvatar(
                                child: Text(
                                  "#${index + 1}",
                                  style: TextStyle(color: Colors.black),
                                ),
                              ),
                              title: buildButton.elementAt(index),
                            ),
                            if (index != buildButton.length - 1)
                              Divider(color: Colors.blue, thickness: 1),
                          ],
                        ),
                        itemCount: buildButton.length,
                      ),
                    ),
                  ),
                  SizedBox(
                    height: 40,
                  ),
                  Center(
                    child: Column(
                      children: [
                        Container(
                          height: 30,
                          child: DefaultTextStyle(
                            style: const TextStyle(
                              fontSize: 20.0,
                              color: Colors.black,
                              fontFamily: 'Canterbury',
                            ),
                            child: Row(
                              mainAxisAlignment: MainAxisAlignment.center,
                              mainAxisSize: MainAxisSize.min,
                              children: [
                                Icon(Icons.person, size: 30),
                                SizedBox(width: 5),
                                AnimatedTextKit(
                                  animatedTexts: [
                                    WavyAnimatedText('Programmed BY:',
                                        textStyle: TextStyle(
                                            fontWeight: FontWeight.bold)),
                                  ],
                                  repeatForever: true,
                                )
                              ],
                            ),
                          ),
                        ),
                        const SizedBox(height: 20),
                        SizedBox(
                          height: 35,
                          width: 300,
                          child: Card(
                            color: HexColor("#6fa8dc"),
                            child: Marquee(
                              text: "Abdelrahman Emad Eldeen Ragab",
                              style: TextStyle(fontWeight: FontWeight.bold),
                              scrollAxis: Axis.horizontal,
                              crossAxisAlignment: CrossAxisAlignment.center,
                              blankSpace: 50,
                              accelerationDuration: Duration(microseconds: 30),
                            ),
                          ),
                        ),
                        const SizedBox(height: 10),
                        SizedBox(
                          height: 35,
                          width: 300,
                          child: Card(
                            color: HexColor("#6fa8dc"),
                            child: Marquee(
                              text: "Mostafa Atta Mohamed",
                              style: TextStyle(fontWeight: FontWeight.bold),
                              scrollAxis: Axis.horizontal,
                              crossAxisAlignment: CrossAxisAlignment.center,
                              blankSpace: 50,
                              accelerationDuration: Duration(microseconds: 30),
                            ),
                          ),
                        ),
                      ],
                    ),
                  ),
                ],
              ),
            ),
          ),
        ),
      ),
    );
  }
}
