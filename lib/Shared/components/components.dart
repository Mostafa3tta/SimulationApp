import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:flutter_animated_button/flutter_animated_button.dart';
import 'package:fluttertoast/fluttertoast.dart';

Widget defaultButton({
  double width = double.infinity,
  Color background = Colors.blue,
  required Function function,
  required String text,
  double redius = 0,
  bool isUpperCase = true,
}) =>
    Container(
      height: 40,
      width: width,
      child: ElevatedButton(
        onPressed: () {
          function();
        },
        child: Text(
          isUpperCase ? text.toUpperCase() : text,
          style: TextStyle(
            color: Colors.white,
          ),
        ),
      ),
      decoration: BoxDecoration(
        borderRadius: BorderRadius.circular(redius),
        color: background,
      ),
    );

Widget defaultFormField({
  required TextEditingController controller,
  required TextInputType type,
  Function? onSubmitted,
  required Function onChange,
  Function? onTap,
  required Function validate,
  required String label,
  String? hint,
  bool? isPassword = false,
  required IconData prefix,
  IconData? suffix,
  Function? suffixPressed,
  bool isClickable = true,
  bool isVaildator = false,
}) {
  return TextFormField(
    controller: controller,
    keyboardType: type,
    obscureText: isPassword!,
    enabled: isClickable,
    onFieldSubmitted: (String? s) {
      onSubmitted!(s);
    },
    onTap: () {
      onTap!();
    },
    onChanged: onChange(),
    validator: (s) {
      if (isVaildator == true) if (s!.isEmpty && s.trim().length == 0) {
        return "${label} mustn't be empty";
      }
      return null;
    },
    decoration: InputDecoration(
      focusedBorder: OutlineInputBorder(
        borderRadius: BorderRadius.circular(25.0),
        borderSide: BorderSide(color: Color.fromRGBO(193, 154, 107, 1)),
      ),
      enabledBorder: OutlineInputBorder(
        borderRadius: BorderRadius.circular(25.0),
        borderSide: BorderSide(
          color: Colors.amber,
          width: 2.0,
        ),
      ),
      labelText: label,
      labelStyle: TextStyle(
        color: Colors.amber,
      ),
      hintText: hint,
      border: OutlineInputBorder(),
      prefixIcon: Icon(
        prefix,
        color: Colors.amberAccent,
      ),
      suffixIcon: suffix != null
          ? IconButton(
              onPressed: () {
                suffixPressed!();
              },
              icon: Icon(suffix),
            )
          : null,
    ),
  );
}

Widget myDivider() {
  return Padding(
    padding: const EdgeInsets.all(8.0),
    child: Container(
      width: double.infinity,
      height: 1.0,
      color: Colors.grey[300],
    ),
  );
}

void navigateTo(context, widget) {
  Navigator.push(
    context,
    MaterialPageRoute(
      builder: (context) => widget,
    ),
  );
}

void navigateAndFinish(context, widget) {
  Navigator.pushAndRemoveUntil(
      context,
      MaterialPageRoute(
        builder: (context) => widget,
      ),
      (route) => false);
}

void defaultToastMsg({
  required String Message,
  required ToastStates state,
}) =>
    Fluttertoast.showToast(
        msg: Message,
        toastLength: Toast.LENGTH_SHORT,
        gravity: ToastGravity.BOTTOM,
        timeInSecForIosWeb: 5,
        backgroundColor: chooseToastColor(state),
        textColor: Colors.white,
        fontSize: 16.0);

enum ToastStates { SUCCESS, ERROR, WARNING }

Color chooseToastColor(ToastStates state) {
  Color color;
  switch (state) {
    case ToastStates.SUCCESS:
      color = Colors.green;
      break;
    case ToastStates.ERROR:
      color = Colors.red;
      break;
    case ToastStates.WARNING:
      color = Colors.amber;
      break;
  }
  return color;
}
