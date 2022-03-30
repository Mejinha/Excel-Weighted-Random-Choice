# Weighted Random Choice Function

## What does this code do?

This code creates a VBA function to select a random element of an array, weighted by an array of weights or frequencies. 

The idea for this code was developed a few years ago while working as an auditor for an insurance multinational. That time, we were highly dependant of bult-in excel solutions for selecting samples. The solution at the time involved using a combination of RANDBETWEEN, VLOOKUP and INDEX. Although this an easy and fast way of selecting samples, it fails at precision.

## Who can benefit from this code?

*Everyone!* 

## How to Install?

* In the tab **Developer**, open the Visual Basid Editor (if the Developer Tab is not visible in your Excel, please check this link: [How to make the Developer tab visible?])
* Go to **Files >> Import Files**
* Import the module **ramdom_choice.bas**

## How to Use?

The usage is very simple. You just need to call the function CHOICE, as shown below:

```sh
=WCHOOSE(Labels, [Weights])
```

The function **WCHOOSE** takes two arguments: Labels and Weights (Optional).

**Labels** is the range of elements that identify each observation to be passed to the function.

**Weights** is the range of weights (frequencies or probabilities) with the same length as Labels. If Weights are not passed to the function, all elements in Labels have the same probability of been chosen.

## Example

|      | A      |  B           |
|------     | ------      | ------      |
|1| **Color**   | **Probability** |
|2| Blue    | 50%     
|3| Red     | 10%
|4| Orange  | 12%
|5| White   | 13%
|6| Black   | 15%


```sh
=WCHOOSE(A2:A6, B2:B6)
```

## Tips

* Remember to "lock" the range of the function before dragging it to multiple lines.
* If you are using probabilities as weights, make sure your values sum up to 1 / 100%.
* The frequencies of the selected elements will eventually approximate to the probabilities of each element if the number of draws is sufficiently large.

[How to make the Developer tab visible?]: <https://support.microsoft.com/en-us/topic/show-the-developer-tab-e1192344-5e56-4d45-931b-e5fd9bea2d45>
