EngCel
======

Engineering extensions for Microsoft Excel

EngCel is an add-in that will ultimately provide extended functionality to make Excel more suitable for Engineering Calculations.

Currently Excel 2003 - 2013 is supported, however it is likely to be compatible with older/newer versions.

A semi-comprehensive illustrated introduction can be found at:
* [In-line super/sub script in Excel and More - Being Brunel](http://www.beingbrunel.com/inline-subsuper-script-in-excel-and-more/ "Being Brunel")
* [EngCel: Colours, Names and More - Being Brunel](http://www.beingbrunel.com/engcel-colours-names-and-more/ "Being Brunel")


Installation
------------

EngCel is a VBA add-in, which means it can be installed from Excel. This can be done for a single session just by opening the file, or permanently by loading the EngCel.xla file.

The method for loading add-ins varies between versions of Excel:

* [Loading an add-in for Excel 2007 to 2013](http://office.microsoft.com/en-us/excel-help/load-or-unload-add-in-programs-HP010096834.aspx#BMexceladdin "Excel 2007 to 2010")
* [Loading an add-in for Excel 2003](http://office.microsoft.com/en-us/excel-help/load-or-unload-add-in-programs-HP005203732.aspx "Excel 2003")


Usage
=====

When loaded EngCel provides the following extensions:

In-line Engineering Notation
----------------------------

The Pipe Parser extension loads automatically with the EngCel. This provides a simple mark-up language, based on LaTeX, which allows super/sub script notation and Greek symbols to be typed directly into the formula bar.

To use the parser start a cell with a pipe "|" this indicates that the text in the cell should be processed using the following rules:

* ^ - The next character should be in super-script
* _ - The next character should be in sub-script
* ¬ - The preceding character should have a bar over it
* {} - Mark-up between brackets should be considered as a single character
* \Alpha - Insert the Greek symbol stated, where proper case denotes a capital, and lower case a little letter. There are also a number of mathematical symbols available, however I haven't had the chance to define these just yet- just give it a go!
* \ - The next character should be escaped if it is a special character

For example:

```
|kN/mm^2
```

Would become kN/mm²

Although why you'd want to write that, I'm really not sure!

Mixins
------

Mixin codes are essentially templates; by piping the name of the mixin the cell will be replaced with a templates cell. At the moment mixins can only be hardcoded into EngCel, however alternatives are being investigated. The following mixins are known:

* |@pass - A 'Pass' or 'Fail', green/red conditionally formatted cell referencing the utilisation percentage two left from the cell. To reverse the sense use @!pass.
* |@ok - A 'OK' or 'Check', green/red conditionally formatted cell referencing the utilisation percentage two left from the cell. To reverse the sense use @!ok.
* |@eng - Format cell to shown engineering exponent notation (i.e. x10^3, ^6, ^-3), etc.
* |@colour - Switch on and off function colouring, see below for more information.
* |@help - Create a new sheet with this help file and an character map

Variable Naming
---------------

The variable namer is an extension to the pipe parser. When invoked it will name the cell to the left of the symbol with the text in the cell to the right of the symbol. Considering that the standard calculation sheet writes [variable name] | [variable symbol] | [variable value] | [units], the aim is to make it easier to name and refer to the variables used.

To invoke the variable namer simply end any pipe notation with '=>'.

For example having the following table:

```
| HAL | |\omega=> | 9000
```

Would result in the cell with '9000' being named 'HAL'. Further formulas can then be more sensical as '=IF(B3=9000,"Can't do that","Can do that")' becomes '=IF(HAL=9000,"Can't do that","Can do that")'.

Note that Excel has some conventions with naming, so stick to unique, character only variable names if you want an easy life.

Function Colouring
------------------

This modification to Excel will colour functions depending on their relationships:

* Formulas that are neither used by, or use other cells will turn grey (unused value)
* Formulas that are used, but not used by, other cells will turn red (input value)
* Formulas that only reference a single cell, without calculation, will turn orange (refereed value)
* Formulas that rely and are relied upon by other cells will turn blue (function)
* Formulas that use other cells but are not used themselves will turn green (result value)

Because of limitations in Excel, mod_colour disables the undo function, and therefore is switched off by default.

To enable mod_colour, use the mixin |@colour. This will return either 'on' or 'off' representing the state of mod_colour after the switch. Note that mod_colour will only effect cells altered when it is on.


Author and Licence
==================

Challenger is primarily written by Thomas Michael Wallace (www.thomasmichaelwallace.co.uk), and released under the GPL v3 licence.
