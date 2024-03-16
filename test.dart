String getColName(String start, int n) {
  int ascii = start.codeUnitAt(0);
  ascii += n;

  if (ascii > 90) {
    int overflow = ascii - 90;
    int firstChar = 64 + (overflow / 26).floor();
    int secondChar = 64 + (overflow % 26);
    if (secondChar == 64) {
      firstChar -= 1;
      secondChar = 90;
    }
    return String.fromCharCode(firstChar+1) + String.fromCharCode(secondChar);
  } else {
    return String.fromCharCode(ascii);
  }
}

void main(List<String> args) {
  // print(getColName("Z", 20));
  for (var i = 0; i < 20; i++) {
    print(getColName("Z", i));
  }
}