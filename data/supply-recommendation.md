# Supply Opportunity Recommendation

Based on our analysis of crossbody sales, we will suffer 350,322.15 lost sales, 
face inventory risk of 642.82, and should order 
0.00 of product to maximize strong sales.

Parameters used: Lead time: 12 weeks

## Lost Sales from Expected Stockouts

| Metric | Value |
|:-------|------:|
| Total Value | 350,322.15 |
| Number of Products | 25 |
| Total Units | 3,047 |


The top 10 lost sales products by value are shown below:

|               Product Description |   ROS |   End Inventory |   Weeks on Hand |   Weekly Sales |   Lost Sales Units |   Lost Sales Value |   Lost Margin |   % of Total Lost Sales |
|----------------------------------:|------:|----------------:|----------------:|---------------:|-------------------:|-------------------:|--------------:|------------------------:|
| Handbag Crossbody bag Style 46648 |  0.29 |             536 |            6.94 |          77.25 |                391 |          54,060.66 |        696.98 |                   15.43 |
| Handbag Crossbody bag Style 53554 |  0.27 |               0 |            0    |          26.75 |                321 |          35,935.26 |     27,024.30 |                   10.26 |
| Handbag Crossbody bag Style 46672 |  0.23 |             474 |            7.8  |          60.75 |                255 |          31,869.14 |     15,398.69 |                    9.1  |
| Handbag Crossbody bag Style 53803 |  0.25 |               0 |            0    |          18.75 |                225 |          27,821.04 |     15,452.79 |                    7.94 |
| Handbag Crossbody bag Style 53562 |  0.28 |               0 |            0    |          23.5  |                282 |          27,013.04 |      8,781.74 |                    7.71 |
| Handbag Crossbody bag Style 43477 |  0.11 |             180 |            5.54 |          32.5  |                210 |          26,682.22 |     19,670.32 |                    7.62 |
| Handbag Crossbody bag Style 53796 |  0.24 |             322 |            7.32 |          44    |                206 |          22,578.62 |     14,215.02 |                    6.45 |
| Handbag Crossbody bag Style 53545 |  0.25 |               0 |            0    |          14.25 |                171 |          18,635.89 |      9,285.61 |                    5.32 |
| Handbag Crossbody bag Style 40077 |  0.18 |             435 |            8.45 |          51.5  |                183 |          14,731.78 |     10,832.05 |                    4.21 |
| Handbag Crossbody bag Style 50012 |  0.12 |             280 |            8.62 |          32.5  |                110 |          14,199.84 |      9,309.24 |                    4.05 |

Methodology: Lost sales analysis focuses on products with above-median ROS (0.094) 
that risk stockout before lead time. For these products, lost sales are calculated based on the gap between 
current weeks of hand and lead time, multiplied by weekly sales rate. Products with zero sales are excluded 
from lost sales calculations.

## Inventory Risk

| Metric | Value |
|:-------|------:|
| Total Value | 642.82 |
| Number of Products | 53 |
| Total Units | 7 |


All 2 overstock products are:

|               Product Description |   ROS |   End Inventory |   Weekly Sales |   Inventory Risk Units |   Inventory Risk Value |   % of Total Risk |
|----------------------------------:|------:|----------------:|---------------:|-----------------------:|-----------------------:|------------------:|
| Handbag Crossbody bag Style 45278 |     0 |               5 |              0 |                      5 |                    485 |             75.41 |
| Handbag Crossbody bag Style 49779 |     0 |               2 |              0 |                      2 |                    158 |             24.59 |

Methodology: Inventory risk assessment identifies excess units beyond lead time plus 4 weeks of cover, based 
on current sales rates. For products with zero sales, all current inventory is considered at risk. Risk value 
is calculated using the product's average unit retail price.

## Restock Opportunities - Green Area

| Metric | Value |
|:-------|------:|
| Total Value | 0.00 |
| Number of Products | 0 |
| Total Units | 0 |

### No restock candidate were identified

Methodology: Restock analysis identifies high-performing products (above median ROS) with inventory levels 
between lead time and lead time plus 4 weeks of cover. Reorder quantities aim to restore stock to lead time 
plus 4 weeks of cover, based on current sales rates. Products with zero sales are excluded from restock 
recommendations.
