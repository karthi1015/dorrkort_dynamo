# Using an excel sheet as dictionary in Dynamo

## Preface
Some of this text might not be interesting to everyone just looking for a solution to import Excel sheets as dictionaries into Dynamo, but for Swedish professionals with the same problem or someone who loves to read, the rest gives you some context.

## The "Dörrkort"-problem
The office I'm working in is switching from a long legacy in in AutoCAD Architecture over to Revit. There are some test projects and I'm part in one of them. The plan is to create a smooth workflow that enables us to take advantage of the built in solutions but also exended functionality of Revit and BIM. 
A very tidious job when building here in Sweden is the so called "Dörrkort". That's basically one door per drawing and you have all door information sorted in a standardized manor. The only good example I could find was in [this master thesis](http://www.diva-portal.se/smash/get/diva2:1221728/FULLTEXT01.pdf) on page 77. Our current workflow in AutoCAD demands a lot of manual checking and filling out text in the respective row. That's just how it was handled so far, there might be a better solution for AutoCAD as well, but my focus is now on Revit.
The biggest part is definitely to define door specific locks and acessories like automatic openers and more. It feels like it can't be so difficult, but someone actually needs to define them and in my current project I have ~180 doors with 27 different lock/acessory-combinations. There are obviously projects with even more doors, that's why I want to automate as much as possible.

## Theory
Lets imagine you have a Excel list from a consultant with different features that apply to more than just one instance in your Revit project. I will stick with my example, 10 doors with 5 different lock types, but each door has a lock meaning that some locks share the same lock type. That's a perfect scenario for a dictionary.
This could be the the excel sheet, in reality there are more columns and more rows. DK is the key for the dictionary.
|DK|Material|Lock|Handle|
|-|-|-|-|
|DK1|Steel|A|round|
|DK2|Steel|B|straight|
|DK3|Steel|C|straight|
|DK4|Aluminum|D|straight|
|DK5|Steel|E|straight|

There are 10 doors in the project:
|ID|DK|Material|Lock|Handle|
|-|-|-|-|-|
|00|DK5|-|-|-|
|01|DK2|-|-|-|
|02|DK3|-|-|-|
|03|DK1|-|-|-|
|04|DK5|-|-|-|
|05|DK4|-|-|-|
|06|DK2|-|-|-|
|07|DK1|-|-|-|
|08|DK1|-|-|-|
|09|DK4|-|-|-|

In the end the list should look like this, again imagine that for 180 doors, more parameters and surely more complex names as well as the possibility to have changes in either list over time, it's a nightmare to do that by hand:
|ID|DK|Material|Lock|Handle|
|-|-|-|-|-|
|00|DK5|Steel|E|straight|
|01|DK2|Steel|B|straight|
|02|DK3|Steel|C|straight|
|03|DK1|Steel|A|round|
|04|DK5|Steel|E|straight|
|05|DK4|Aluminum|D|straight|
|06|DK2|Steel|B|straight|
|07|DK1|Steel|A|round|
|08|DK1|Steel|A|round|
|09|DK4|Aluminum|D|straight|

## What Dynamo does
Dynamo can't apply a dictionary from an excel sheet and just match the key value of the excel sheet with the key value of Revit objects out of the box (afaik). You basically have to tell Dynamo to get all the doors, and then you match each parameter in the door to the parameter in the excel sheet/dictionary. That way you get a long list. I would say it's a workaround, there should be a node that does that out of the box.

Here is what I did:
![Dynamo Script](https://aws1.discourse-cdn.com/business6/uploads/dynamobim/original/3X/f/5/f53ad7040bbd17142d64fe8fb7dcf2c8cb4d1d09.png)

## Preparations
To have the script working you need to create parameters in your project. I used shared parameters that I loaded as project parameters and in this case most of them are instance-based. But how you handle your parameters is your thing, the important part is that this script doesn't create the parameters, you have to have them beforehand! If you google a bit you'll find tips how to batch create parameters.

## Walk through the script
### Step 1 - loading Revit elements
![Step 1](https://github.com/monsieurhannes/dorrkort_dynamo/blob/main/screenshots/step_1.png)
That's quite basic, but worth mentioning when using Dynamo for the first time. You can use the script for any other category.

### Step 2 - loading and processing the Excel sheet
![Step 2](https://github.com/monsieurhannes/dorrkort_dynamo/blob/main/screenshots/step_2.png)
Excel import has it's own node, that's quite ease, then I use some Python scripts involved, it might be possible without python, but it's faster for me to write a python script. Let's follow the upper line first. It takes the headers of the Excel sheet, Those are the parameter names that get filled out later - it's important that there are no typos! Then I drop the key row as I only want to have the parameters that I want to add to the doors, the key stays unchanged.

