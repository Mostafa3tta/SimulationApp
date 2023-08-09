import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:simulation_excel/stayles/colors.dart';
import 'package:hexcolor/hexcolor.dart';


ThemeData lightMode = ThemeData(
  textTheme: TextTheme(
    bodyText1: TextStyle(
      fontSize: 16,
      fontWeight: FontWeight.w600,
      color: Colors.black,
      fontFamily: 'Jannah',
    ),
    bodyText2: TextStyle(
        fontSize: 18.0,
        color: Colors.grey[800],
        fontFamily: 'Jannah',

    ),
    subtitle1: TextStyle(
      fontSize: 15.0,
      color: Colors.black,
      fontFamily: 'Jannah',
      height: 1.4,
    ),
  ),
  primarySwatch: defaultColor,
  scaffoldBackgroundColor: Colors.white,
  floatingActionButtonTheme: FloatingActionButtonThemeData(
    backgroundColor: defaultColor,
    splashColor: Colors.orange,
  ),
  appBarTheme: AppBarTheme(
      titleSpacing: 20.0,
      backwardsCompatibility: false,
      // False: بيخليني اتحكم في النوتفكيشن بار
      systemOverlayStyle: SystemUiOverlayStyle(
        statusBarColor: Colors.white,
        statusBarIconBrightness: Brightness.dark,
      ),
      titleTextStyle: TextStyle(
        color: Colors.black,
        fontSize: 20.0,
        fontWeight: FontWeight.bold,
      ),
      backgroundColor: Colors.white,
      elevation: 0.0,
      iconTheme: IconThemeData(
        color: defaultColor,
      )),
  bottomNavigationBarTheme: BottomNavigationBarThemeData(
    type: BottomNavigationBarType.fixed,
    selectedItemColor: defaultColor,
    elevation: 20.0,
  ),
);

ThemeData darkMode = ThemeData(
  cardColor: Colors.black,
  primarySwatch: defaultColor,
  textTheme: TextTheme(
    bodyText1: TextStyle(
      fontSize: 16,
      fontWeight: FontWeight.w600,
      color: Colors.white,
      fontFamily: 'Jannah',
    ),
    bodyText2: TextStyle(
        fontSize: 15.0,
        color: Colors.white,
        fontFamily: 'Jannah',
    ),
    subtitle1: TextStyle(
      fontSize: 15.0,
      color: Colors.white,
      fontFamily: 'Jannah',
      height: 1.3
    ),
  ),
  scaffoldBackgroundColor: Colors.black,//HexColor('333739'),
  appBarTheme: AppBarTheme(
      titleSpacing: 20.0,
      backwardsCompatibility: false,
      // False: بيخليني اتحكم في النوتفكيشن بار
      systemOverlayStyle: SystemUiOverlayStyle(
        statusBarColor: Colors.black ,//HexColor('333739'),
        statusBarIconBrightness: Brightness.light,
      ),
      titleTextStyle: TextStyle(
        color: Colors.white,
        fontSize: 20.0,
        fontWeight: FontWeight.bold,
      ),
      backgroundColor: Colors.black,//HexColor('333739'),
      elevation: 0.0,
      iconTheme: IconThemeData(
        color: Colors.white,
      )),
  bottomNavigationBarTheme: BottomNavigationBarThemeData(
    backgroundColor: HexColor('333739'),
    unselectedItemColor: Colors.white,
    type: BottomNavigationBarType.fixed,
    selectedItemColor: defaultColor,
    elevation: 20.0,
  ),
);


