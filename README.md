# PCVolumeMixer

This was a very rough Proof of Concept I made back in 2018 and recently came across the code for. It is a volume mixer for Windows that lets you control the volume of individual applications on a PC.

It uses an Arduino Pro Micro, a custom PCB, and four potentiometers. The Pro Micro communicates with the C# application over a USB serial connection.
The pots are connected to A0-A3 on the Pro Micro.

<p align="center">
<img alt="3D Printed Case" src="https://github.com/blaketilghman/PCVolumeMixer/blob/master/images/Case.PNG?raw=true" width="200"></br>
<img alt="Custom PCB" src="https://github.com/blaketilghman/PCVolumeMixer/blob/master/images/PCB.PNG?raw=true" width="200"></br>
<img alt="Screenshot" src="https://github.com/blaketilghman/PCVolumeMixer/blob/master/images/Capture.PNG?raw=true" width="200"></br>
</p>

Included are the 3D model files that I made for the case, the gerber files for the PCB, and all the code that I could find.

All code was written on Visual Studio Community 2017 and has not been tested on any newer versions.
