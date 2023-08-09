import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:simulation_excel/stayles/icon_brocken.dart';
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
  required Function onTap,
  required Function validate,
  required String label,
  String? hint,
  bool? isPassword = false,
  required IconData prefix,
  IconData? suffix,
  Function? suffixPressed,
  bool isClickable = true,
  bool isVaildator = false,
  double ? focusedBorderReduis ,
  double ? enabledBorderReduis  ,
}) {
  return TextFormField(
    controller: controller,
    keyboardType: type,
    obscureText: isPassword!,
    enabled: isClickable,
    onFieldSubmitted: (String? s) {
      onSubmitted!(s);
    },
    onTap: onTap(),
    onChanged: onChange(),
    validator: (s) {
      if (isVaildator == true) if (s!.isEmpty && s.trim().length == 0) {
        return "${label} mustn't be empty";
      }
      return null;
    },
    decoration: InputDecoration(
      labelText: label,
      hintText: hint,
      focusedBorder: OutlineInputBorder(
        borderRadius: BorderRadius.circular(25.0),
        borderSide: BorderSide(
        ),
      ),
      enabledBorder: OutlineInputBorder(
        borderRadius: BorderRadius.circular(25.0),
        borderSide: BorderSide(
        ),
      ),
      border: OutlineInputBorder(),
      prefixIcon: Icon(
        prefix,
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

PreferredSizeWidget defaultAppBar(
{
  required BuildContext context,
  String ?title,
  List<Widget>? action,
}
    ){
  return  AppBar(
    leading: IconButton(
      onPressed: (){
        Navigator.pop(context);
      },
      icon: Icon(IconBroken.Arrow___Left_2),
    ),
    title: Text(title!),
    titleSpacing: 5.0,
    actions: action,
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

