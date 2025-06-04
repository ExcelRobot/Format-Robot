# <h1 id="oa-robot-definitions">OA Robot Definitions</h1>

\*\*Format Robot Vol 1.xlsm\*\* contains definitions for:

[26 Robot Commands](#command-definitions)<BR>[2 Robot Texts](#text-definitions)<BR>

<BR>

## Available Robot Commands

[Fill](#fill) | [Font](#font) | [Number](#number)

### Fill

| Name | Description |
| --- | --- |
| [Format Fill Light Blue](#format-fill-light-blue) | Format selection with light blue fill. |
| [Format Fill Light Green](#format-fill-light-green) | Format selection with light green fill. |
| [Format Fill Light Orange](#format-fill-light-orange) | Format selection with light orange fill. |
| [Format Fill Light Purple](#format-fill-light-purple) | Format selection with light purple fill. |
| [Format Fill Light Red](#format-fill-light-red) | Format selection with light red fill. |
| [Format Fill Light Yellow](#format-fill-light-yellow) | Format selection with light yellow fill. |
| [Remove Fill Colors](#remove-fill-colors) | Removes all fill colors from selection. |

### Font

| Name | Description |
| --- | --- |
| [Format Font Black](#format-font-black) | Format selection with black font. |
| [Format Font Blue](#format-font-blue) | Format selection with blue font. |
| [Format Font Grey](#format-font-grey) | Format selection with grey font. |
| [Format Font Light Grey](#format-font-light-grey) | Format selection with light grey font. |
| [Format Font Red](#format-font-red) | Format selection with red font. |
| [Format Font White](#format-font-white) | Format selection with white font. |

### Number

| Name | Description |
| --- | --- |
| [Format Accounting 0 Decimals](#format-accounting-0-decimals) | Format selection as Accounting with no decimals. |
| [Format Accounting 1 Decimal](#format-accounting-1-decimal) | Format selection as Accounting with one decimal. |
| [Format Accounting 2 Decimals](#format-accounting-2-decimals) | Format selection as Accounting with two decimals. |
| [Format Multiple 1 Decimal](#format-multiple-1-decimal) | Format selection as multiple with one decimal (0.0x). |
| [Format Multiple 2 Decimals](#format-multiple-2-decimals) | Format selection as multiple with two decimals (0.00x). |
| [Format Number 0 Decimals](#format-number-0-decimals) | Format selection as Number with no decimals. |
| [Format Number 1 Decimal](#format-number-1-decimal) | Format selection as Number with one decimal. |
| [Format Number 2 Decimals](#format-number-2-decimals) | Format selection as Number with two decimals. |
| [Format Number 3 Decimals](#format-number-3-decimals) | Format selection as Number with three decimals. |
| [Format Percent 0 Decimals](#format-percent-0-decimals) | Format selection as Percent with no decimals. |
| [Format Percent 1 Decimal](#format-percent-1-decimal) | Format selection as Percent with one decimal. |
| [Format Percent 2 Decimals](#format-percent-2-decimals) | Format selection as Percent with two decimals. |
| [Format Short Date](#format-short-date) | Format selection as short date (m\/d\/yyyy). |

<BR>

## Available Robot Texts

| Name | Description |
| --- | --- |
| [Format Robot Volume 1 Description](#format-robot-volume-1-description) | |
| [Format Robot Volume 1 Examples](#format-robot-volume-1-examples) | |

<BR>

## Command Definitions

<BR>

### Format Accounting 0 Decimals

*Format selection as Accounting with no decimals.*

<sup>`@Format Robot Vol 1.xlsm` `!VBA Macro Command` `#Number`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modFormat.ApplyNumberFormat](./VBA/modFormat.bas#L12)([Selection],"_(* #,##0_);_(* (#,##0);_(* "" - ""??_);_(@_)")</code> |
| Launch Codes | <code>a0</code> |

[^Top](#oa-robot-definitions)

<BR>

### Format Accounting 1 Decimal

*Format selection as Accounting with one decimal.*

<sup>`@Format Robot Vol 1.xlsm` `!VBA Macro Command` `#Number`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modFormat.ApplyNumberFormat](./VBA/modFormat.bas#L12)([Selection],"_(* #,##0.0_);_(* (#,##0.0);_(* "" - ""??_);_(@_)")</code> |
| Launch Codes | <code>a1</code> |

[^Top](#oa-robot-definitions)

<BR>

### Format Accounting 2 Decimals

*Format selection as Accounting with two decimals.*

<sup>`@Format Robot Vol 1.xlsm` `!VBA Macro Command` `#Number`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modFormat.ApplyNumberFormat](./VBA/modFormat.bas#L12)([Selection],"_(* #,##0.00_);_(* (#,##0.00);_(* "" - ""??_);_(@_)")</code> |
| Launch Codes | <code>a2</code> |

[^Top](#oa-robot-definitions)

<BR>

### Format Fill Light Blue

*Format selection with light blue fill.*

<sup>`@Format Robot Vol 1.xlsm` `!VBA Macro Command` `#Fill`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modFormat.ApplyBackgroundColorRGB](./VBA/modFormat.bas#L20)([Selection],218,233,248)</code> |

[^Top](#oa-robot-definitions)

<BR>

### Format Fill Light Green

*Format selection with light green fill.*

<sup>`@Format Robot Vol 1.xlsm` `!VBA Macro Command` `#Fill`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modFormat.ApplyBackgroundColorRGB](./VBA/modFormat.bas#L20)([Selection],218,242,208)</code> |

[^Top](#oa-robot-definitions)

<BR>

### Format Fill Light Orange

*Format selection with light orange fill.*

<sup>`@Format Robot Vol 1.xlsm` `!VBA Macro Command` `#Fill`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modFormat.ApplyBackgroundColorRGB](./VBA/modFormat.bas#L20)([Selection],251,226,213)</code> |

[^Top](#oa-robot-definitions)

<BR>

### Format Fill Light Purple

*Format selection with light purple fill.*

<sup>`@Format Robot Vol 1.xlsm` `!VBA Macro Command` `#Fill`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modFormat.ApplyBackgroundColorRGB](./VBA/modFormat.bas#L20)([Selection],242,206,239)</code> |

[^Top](#oa-robot-definitions)

<BR>

### Format Fill Light Red

*Format selection with light red fill.*

<sup>`@Format Robot Vol 1.xlsm` `!VBA Macro Command` `#Fill`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modFormat.ApplyBackgroundColorRGB](./VBA/modFormat.bas#L20)([Selection],255,0,0,0.8)</code> |

[^Top](#oa-robot-definitions)

<BR>

### Format Fill Light Yellow

*Format selection with light yellow fill.*

<sup>`@Format Robot Vol 1.xlsm` `!VBA Macro Command` `#Fill`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modFormat.ApplyBackgroundColorRGB](./VBA/modFormat.bas#L20)([Selection],255,255,0,0.8)</code> |

[^Top](#oa-robot-definitions)

<BR>

### Format Font Black

*Format selection with black font.*

<sup>`@Format Robot Vol 1.xlsm` `!VBA Macro Command` `#Font`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modFormat.ApplyFontColorRGB](./VBA/modFormat.bas#L16)([Selection],0,0,0)</code> |

[^Top](#oa-robot-definitions)

<BR>

### Format Font Blue

*Format selection with blue font.*

<sup>`@Format Robot Vol 1.xlsm` `!VBA Macro Command` `#Font`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modFormat.ApplyFontColorRGB](./VBA/modFormat.bas#L16)([Selection],0,0,255)</code> |

[^Top](#oa-robot-definitions)

<BR>

### Format Font Grey

*Format selection with grey font.*

<sup>`@Format Robot Vol 1.xlsm` `!VBA Macro Command` `#Font`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modFormat.ApplyFontColorRGB](./VBA/modFormat.bas#L16)([Selection],128,128,128)</code> |

[^Top](#oa-robot-definitions)

<BR>

### Format Font Light Grey

*Format selection with light grey font.*

<sup>`@Format Robot Vol 1.xlsm` `!VBA Macro Command` `#Font`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modFormat.ApplyFontColorRGB](./VBA/modFormat.bas#L16)([Selection],208,208,208)</code> |

[^Top](#oa-robot-definitions)

<BR>

### Format Font Red

*Format selection with red font.*

<sup>`@Format Robot Vol 1.xlsm` `!VBA Macro Command` `#Font`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modFormat.ApplyFontColorRGB](./VBA/modFormat.bas#L16)([Selection],255,0,0)</code> |

[^Top](#oa-robot-definitions)

<BR>

### Format Font White

*Format selection with white font.*

<sup>`@Format Robot Vol 1.xlsm` `!VBA Macro Command` `#Font`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modFormat.ApplyFontColorRGB](./VBA/modFormat.bas#L16)([Selection],255,255,255)</code> |

[^Top](#oa-robot-definitions)

<BR>

### Format Multiple 1 Decimal

*Format selection as multiple with one decimal (0.0x).*

<sup>`@Format Robot Vol 1.xlsm` `!VBA Macro Command` `#Number`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modFormat.ApplyNumberFormat](./VBA/modFormat.bas#L12)([Selection],"#,##0.0x")</code> |
| Launch Codes | <code>m1</code> |

[^Top](#oa-robot-definitions)

<BR>

### Format Multiple 2 Decimals

*Format selection as multiple with two decimals (0.00x).*

<sup>`@Format Robot Vol 1.xlsm` `!VBA Macro Command` `#Number`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modFormat.ApplyNumberFormat](./VBA/modFormat.bas#L12)([Selection],"#,##0.00x")</code> |
| Launch Codes | <code>m2</code> |

[^Top](#oa-robot-definitions)

<BR>

### Format Number 0 Decimals

*Format selection as Number with no decimals.*

<sup>`@Format Robot Vol 1.xlsm` `!VBA Macro Command` `#Number`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modFormat.ApplyNumberFormat](./VBA/modFormat.bas#L12)([Selection],"#,##0")</code> |
| Launch Codes | <code>n0</code> |

[^Top](#oa-robot-definitions)

<BR>

### Format Number 1 Decimal

*Format selection as Number with one decimal.*

<sup>`@Format Robot Vol 1.xlsm` `!VBA Macro Command` `#Number`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modFormat.ApplyNumberFormat](./VBA/modFormat.bas#L12)([Selection],"#,##0.0")</code> |
| Launch Codes | <code>n1</code> |

[^Top](#oa-robot-definitions)

<BR>

### Format Number 2 Decimals

*Format selection as Number with two decimals.*

<sup>`@Format Robot Vol 1.xlsm` `!VBA Macro Command` `#Number`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modFormat.ApplyNumberFormat](./VBA/modFormat.bas#L12)([Selection],"#,##0.00")</code> |
| Launch Codes | <code>n2</code> |

[^Top](#oa-robot-definitions)

<BR>

### Format Number 3 Decimals

*Format selection as Number with three decimals.*

<sup>`@Format Robot Vol 1.xlsm` `!VBA Macro Command` `#Number`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modFormat.ApplyNumberFormat](./VBA/modFormat.bas#L12)([Selection],"#,##0.000")</code> |
| Launch Codes | <code>n3</code> |

[^Top](#oa-robot-definitions)

<BR>

### Format Percent 0 Decimals

*Format selection as Percent with no decimals.*

<sup>`@Format Robot Vol 1.xlsm` `!VBA Macro Command` `#Number`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modFormat.ApplyNumberFormat](./VBA/modFormat.bas#L12)([Selection],"0%")</code> |
| Launch Codes | <code>p0</code> |

[^Top](#oa-robot-definitions)

<BR>

### Format Percent 1 Decimal

*Format selection as Percent with one decimal.*

<sup>`@Format Robot Vol 1.xlsm` `!VBA Macro Command` `#Number`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modFormat.ApplyNumberFormat](./VBA/modFormat.bas#L12)([Selection],"0.0%")</code> |
| Launch Codes | <code>p1</code> |

[^Top](#oa-robot-definitions)

<BR>

### Format Percent 2 Decimals

*Format selection as Percent with two decimals.*

<sup>`@Format Robot Vol 1.xlsm` `!VBA Macro Command` `#Number`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modFormat.ApplyNumberFormat](./VBA/modFormat.bas#L12)([Selection],"0.00%")</code> |
| Launch Codes | <code>p2</code> |

[^Top](#oa-robot-definitions)

<BR>

### Format Short Date

*Format selection as short date (m\/d\/yyyy).*

<sup>`@Format Robot Vol 1.xlsm` `!VBA Macro Command` `#Number`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modFormat.ApplyNumberFormat](./VBA/modFormat.bas#L12)([Selection],"m/d/yyyy")</code> |

[^Top](#oa-robot-definitions)

<BR>

### Remove Fill Colors

*Removes all fill colors from selection.*

<sup>`@Format Robot Vol 1.xlsm` `!VBA Macro Command` `#Fill`</sup>

| Property | Value |
| --- | --- |
| Macro Expression | <code>[modFormat.RemoveBackgroundFillColor](./VBA/modFormat.bas#L4)</code> |

[^Top](#oa-robot-definitions)

<BR>

## Text Definitions

<BR>

### Format Robot Volume 1 Description

<sup>`@Format Robot Vol 1.xlsm` `!Default Text` </sup>

| Property | Value |
| --- | --- |
| Text | [Format Robot Volume 1 Description](<./Text/Format Robot Volume 1 Description.txt>) |
| Value | <code>\# Format Robot Volume 1</code><br><code></code><br><code>\#\# Overview</code><br><code></code><br><code>Format Robot Volume 1 is a collection of commands that apply simple number, text, or fill formatting to the current selection.</code><br><code></code><br><code>\#\# Number Formatting</code><br><code></code><br><code>\*\*Accounting:\*\* Excel's standard Accounting format with no currency symbol. Negatives are in parentheses, zeros are displayed as "\-", and all cells are justified to ali... |

[^Top](#oa-robot-definitions)

<BR>

### Format Robot Volume 1 Examples

<sup>`@Format Robot Vol 1.xlsm` `!Default Text` </sup>

| Property | Value |
| --- | --- |
| Text | [Format Robot Volume 1 Examples](<./Text/Format Robot Volume 1 Examples.txt>) |
| Value | <code>iVBORw0KGgoAAAANSUhEUgAAAbcAAAJ\/CAYAAAD7xJ2vAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAGZYSURBVHhe7b1Prusw0rd31+UFeR29BM96FT31AnocBD3JixefJ42MOuhBXmTSATI4qSqSEouiKMmWLYvneQDee2RSxX+l+kmyTP35AQAA6AzEDQAAugNxAwCA7kDcAACgOxA3AADoDsQNAAC6A3EDAIDuQNwAAKA7EDcAAOgOxA0AALoDcQMAgO5A3AAAoDsQNwAA6A7EDQAAuuNFcXv83C5\/fv5cbvLXOXncbz\/X+\/6tf9yuP5c\/MjaSLtd7\/LQkjl8sZ+ly\/bnt1J539S0Q2j7btRPwuF2m7X\/cZN4uP7fmsD1+7tfLMGeXq\/f\/Vl5C6\/7z5\/ozP3z3n2u0kVLe1sd99C\/1mXya5+vXOVvqG5yGr... |

[^Top](#oa-robot-definitions)
