// FastLED - Version: Latest 
#include <FastLED.h>

// converts the position of a 10k lin(B) pot to 0-100%
// pot connected to A0, 5volt and ground

#define LED_PIN
#define NUM_LEDS

int rawValue0;
int oldValue0;
byte potPercentage0;
byte oldPercentage0;
int rawValue1;
int oldValue1;
byte potPercentage1;
byte oldPercentage1;
int rawValue2;
int oldValue2;
byte potPercentage2;
byte oldPercentage2;
int rawValue3;
int oldValue3;
byte potPercentage3;
byte oldPercentage3;

int b1State = 0;
int b2State = 0;
int b3State = 0;
int b4State = 0;

void setup() {
  pinMode(7,INPUT_PULLUP);
  pinMode(8,INPUT_PULLUP);
  pinMode(9,INPUT_PULLUP);
  pinMode(10,INPUT_PULLUP);
  Serial.begin(115200); // set serial monitor to this baud rate, or change the value
  delay(200);
}

void loop() {
  // read input twice
  rawValue0 = analogRead(A0);
  rawValue0 = analogRead(A0); // double read
  rawValue1 = analogRead(A1);
  rawValue1 = analogRead(A1); // double read
  rawValue2 = analogRead(A2);
  rawValue2 = analogRead(A2); // double read
  rawValue3 = analogRead(A3);
  rawValue3 = analogRead(A3); // double read
  // ignore bad hop-on region of a pot by removing 8 values at both extremes
  rawValue0 = constrain(rawValue0, 8, 1015);
  rawValue1 = constrain(rawValue1, 8, 1015);
  rawValue2 = constrain(rawValue2, 8, 1015);
  rawValue3 = constrain(rawValue3, 8, 1015);
  
  
  b1State = digitalRead(7);
  b2State = digitalRead(8);
  b3State = digitalRead(9);
  b4State = digitalRead(10);
  
  if (b1State == HIGH){
    Serial.println("B0:");
  }
  if (b2State == HIGH){
    Serial.println("B1:");
  }
  if (b3State == HIGH){
    Serial.println("B2:");
  }
  if (b4State == HIGH){
    Serial.println("B3:");
  }
  
  
  //pot1
  // add some deadband
  if (rawValue0 < (oldValue0) || rawValue0 > (oldValue0)) {
  //if (rawValue0 < (oldValue0 - 4) || rawValue0 > (oldValue0 + 4)) {
    oldValue0 = rawValue0;
    
    // convert to percentage
    potPercentage0 = map(oldValue0, 8, 1008, 0, 100);
    
    // Only print if %value changes
    if (oldPercentage0 != potPercentage0) {
    //if (oldPercentage0 == potPercentage0 + 3 || oldPercentage0 == potPercentage0 - 3) {
      Serial.print("A0:");
      Serial.println(potPercentage0);
      oldPercentage0 = potPercentage0;
    }
  }
  
  //pot2
  if (rawValue1 < (oldValue1) || rawValue1 > (oldValue1)) {
  //if (rawValue1 < (oldValue1 - 4) || rawValue1 > (oldValue1 + 4)) {
    oldValue1 = rawValue1;
    
    // convert to percentage
    potPercentage1 = map(oldValue1, 8, 1008, 0, 100);
    
    // Only print if %value changes
    if (oldPercentage1 != potPercentage1) {
    //if (oldPercentage1 == potPercentage1 + 3 || oldPercentage1 == potPercentage1 - 3) {
      Serial.print("A1:");
      Serial.println(potPercentage1);
      oldPercentage1 = potPercentage1;
    }
  }
  
  //pot3
  if (rawValue2 < (oldValue2) || rawValue2 > (oldValue2)) {
  //if (rawValue2 < (oldValue2 - 4) || rawValue2 > (oldValue2 + 4)) {
    oldValue2 = rawValue2;
    
    // convert to percentage
    potPercentage2 = map(oldValue2, 8, 1008, 0, 100);
    
    // Only print if %value changes
    if (oldPercentage2 != potPercentage2) {
    //if (oldPercentage2 == potPercentage2 + 3 || oldPercentage2 == potPercentage2 - 3) {
      Serial.print("A2:");
      Serial.println(potPercentage2);
      oldPercentage2 = potPercentage2;
    }
  }
  
  //pot4
  if (rawValue3 < (oldValue3) || rawValue3 > (oldValue3)) {
  //if (rawValue3 < (oldValue3 - 4) || rawValue3 > (oldValue3 + 4)) {
    oldValue3 = rawValue3;
    
    // convert to percentage
    potPercentage3 = map(oldValue3, 8, 1008, 0, 100);
    
    // Only print if %value changes
    if (oldPercentage3 != potPercentage3) {
    //if (oldPercentage3 == potPercentage3 + 3 || oldPercentage3 == potPercentage3 - 3) {
      Serial.print("A3:");
      Serial.println(potPercentage3);
      oldPercentage3 = potPercentage3;
    }
  }
  
  //Serial.print("A0:");
  //Serial.println(potPercentage0);
  //Serial.print("A1:");
  //Serial.println(potPercentage1);
  //Serial.print("A2:");
  //Serial.println(potPercentage2);
  //Serial.print("A3:");
  //Serial.println(potPercentage3);
  
}
