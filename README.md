# Fractal-Noise
Learning how fractal noise maps work

## Description
Been exploring procedural generation topics for game development so decided to see what I could do in excel to mimic some of the terrain generation techniques used.

## How it works
- Created a noise map array using a RandBetween function with **_Max Height_** and **_Min Height_** variables.
- Apply smoothing recursively to the noise map by getting the average of each "cell" and the "cells" in each cardinal direction for a certain **_depth_**
- Repeat the above 2 steps to a specified number of **_layers_** while applying a multiplier to the **_depth_** number
- Add the **_layers_** together to get the final result

The depth total number is the total number of smoothing that has been applied which works out to be **_depth * multiplier^layers-1_**

The lenth of time it takes to run the higher multiplier items is considerable

| Max Height |  Min Height | Depth | Multiplier | Layers | Depth Total | Result |
| ---------- | ----------- | ----- | ---------- | ------ | ----------- | ------ |
|1	|-1|	2|	3	|5	|162| ![image](https://github.com/tgorman31/Fractal-Noise/assets/47192981/56c23a05-04f1-4ca0-918b-8824690e0224) |
|1	|-1|	2|	2	|5	|32| ![image](https://github.com/tgorman31/Fractal-Noise/assets/47192981/d806f1dd-8732-44f3-8a0a-06af2fee9171)|
|1	|-1|	5|	2	|5	|80| ![image](https://github.com/tgorman31/Fractal-Noise/assets/47192981/59f64ca7-5f0d-4671-8cb5-c96750e18c65)|
|1	|0|	3|	2	|10	|1536|![image](https://github.com/tgorman31/Fractal-Noise/assets/47192981/3772f26b-dc34-4d33-a946-d7ba81bab608)|
|1	|-1|	3|	3	|5	|243|![image](https://github.com/tgorman31/Fractal-Noise/assets/47192981/0299075e-104e-465a-8181-c564bdecbbc4) ![image](https://github.com/tgorman31/Fractal-Noise/assets/47192981/f403b411-2f9c-43fb-b5d2-90815983be1d)|
|1	|0|	3|	3	|6	|729||
