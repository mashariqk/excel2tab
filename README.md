# excel2tab

Convert excel file into a tab delimited file using Apache poi library in java
This is for the scenario where a user uploads an excel file containing multiple orders and they have to be parsed and converted into a tab delimited txt file for eventual upload to SAP (or any other ERP for that matter)

## Code layout

This is a very basic working code to parse an excel file using [Apache poi](https://poi.apache.org/). As such, it is pretty messy in terms of design. This is not intended to be a code that you take and deploy in PROD. However it can be used as a good starting point. A lot of validation and formatting may have to be added before this code becomes deployable.

### New To Maven?

I have created this as a [Maven](https://maven.apache.org/) project since it is so much more easier to maintain .jar files with it. However if you are new to maven and not sure how to begin, then just checkout this repository to your machine and import this in eclipse as an existing maven project. Thats it!
