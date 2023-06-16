# R Syntax for Nelson 1994 T-Chart for Adverse Rare Events.
#
This project developed a T-Chart - a type of control chart used to track the times between adverse rare events - based on formulae by Nelson which were developed in 1994. Some R packages already exist which allow you to produce a variety of control charts including: ggQC, qcc, and qicharts2. However, only qicharts2 allows you to produce T-Charts, and the package is limited in terms of the data it will accept for a T-Chart and the extent to which you can customise them. Specifically, the package does not accept data where there is more than one adverse event on the same day, which creates a time-between of 0 days between these two or more episodes (i.e. zeros are not accepted). 

The formulae can be found in Lloyd Nelson's original paper which explained how to calculate the Centreline, Upper Control Limit, and Lower Control Limit for a T-Chart [^1]. This is the earliest paper identified which explains the need for a T-Chart and how to create one. There is also a slight addition to the original formulae devised by Nelson, so that outliers can be identified and removed prior to calculating the three lines, and this was identified from the textbook by Provost & Murray (2022) [^2].



#
#
## The Files for This Project:

1_R Syntax for Nelson 1994 T-Chart - the R syntax needed to carry out the calculations that underpin the T-Chart and then produce the chart itself.

2_Nelson_1994_T_Chart_Sample_Data - an Excel file with two worksheets, each with a single column of sample data, that can be used to produce a T-Chart.

3_T-Chart Example - a PNG File that shows what the T-Chart looks like using this R syntax.
#
#
## An Example of a T-Chart

An example showing what the T-Chart will look like if plotted using this R syntax.



![3_T-Chart Example](https://github.com/Epidemiology22/R-Syntax-for--Nelson-1994-T-Chart/assets/129181130/c4290c92-d27e-450f-b49f-c688962e251c)

An example showing the T-Chart with data points above the Centreline (green) and below the Centreline (orange) in colour.



![image](https://github.com/Epidemiology22/R-Syntax-for--Nelson-1994-T-Chart/assets/129181130/571912d3-84fd-40bc-97e6-98142e273d0f)
#
#
## References
[^1]: Nelson LS. A Control Chart for Parts-Per-Million Nonconforming Items. Journal of Quality Technology. 1994 Jul 1;26(3):239-40.
[^2]: Provost LP and Murray SK. The Health Care Data Guide: Learning From Data for Improvement. Jossey-Bass. Second Edition. 8th August 2022. Pp. 143, 202-6.


