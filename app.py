import streamlit as st
import pandas as pd
import os
import io
import calendar
import traceback
import numpy as np
from datetime import datetime

# Embedded MTC Fleet Capacity Lookup (3,390 vehicles from MTC _Fleet_List_08012025.xlsx - Sheet3)
FLEET_CAPACITY_GROUPS = {
    25: (
        "J0250 J0251 J0252 J0253 J0254 J0255 J0256 J0257 J0258 J0259 J0261 J0262 J0263 J0264 J0265 J0266 "
        "J0267 J0268 J0269 J0271 J0272 J0273 J0275 J0279 J0280 J0283 J0284 J0285 J0286 J0287 J0288 J0289 "
        "J0290 J0293 J0294 J0295 J0296 J0299 J0301 J0303 J0304 J0305 J0306 J0307 J0308 J0309 J0310 J0311 "
        "J0312 J0313 J0314 J0315 J0316 J0317 J0319 J0320 J0321 J0322 J0323 J0324 J0325 J0326 J0327 J0328 "
        "J0329 J0330 J0332 J0334 J0335 J0336 J0337 J0338 J0339 J0340 J0341 J0342 J0343 J0344 J0346 J0347 "
        "J0348 J0349 J0600 J0601 J0602 J0603 J0604 J0605 J0606 J0607 J0608 J0609 J0610 J0611 J0612 J0613 "
        "J0614 J0615 J0616 J0617 J0618 J0619 J0620 J0621 J0622 J0623 J0624 J0625 J0626 J0628 J0629 J0630 "
        "J0631 J0632 J0633 J0634 J0635 J0636 J0637 J0638 J0639 J0640 J0641 J0642 J0643 J0646 J0647 J0648 "
        "J0649 J0650 J0651 J0652 J0653 J0654 J0655 J0656 J0657 J0658 J0659 J0660 J0661 J0662 J0663 J0664 "
        "J0665 J0666 J0668 J0669 J0670 J0671 J0672 J0673 J0674 J0675 J0676 J0677 J0678 J0679 J0680 J0681 "
        "J0682 J0683 J0684 J0686 J0687 J0688 J0689 J0690 J0691 J0692 J0693 J0694 J0695 J0696 J0697 J0698 "
        "J0699 J0700 J0701 J0702 J0703 J0705 J0706 J0707 J0709 J1049 J1050 J0644 J0645 J0685 J0708 S0008 "
        "S0007 "
    ),
    60: (
        "J0457 J0459 J0460 J0461 J0464 J0467 J0468 J0469 J0472 J0473 J0480 J0481 J0482 J0484 J0486 J0487 "
        "J0488 J0489 J0490 J0491 J0498 J0499 J0500 J0514 J0522 J0524 J0533 J0534 J0537 J0538 J0539 J0540 "
        "J0542 J0543 J0544 J0546 J0547 J0549 J0570 J0761 J0762 J0763 J0765 J0766 J0767 J0768 J0769 J0770 "
        "J0771 J0773 J0774 J0775 J0776 J0777 J0778 J0782 J0783 J0784 J0785 J0786 J0787 J0788 J0789 J0790 "
        "J0791 J0792 J0793 J0794 J0795 J0796 J0797 J0800 J0801 J0803 J0804 J0805 J0806 J0807 J0808 J0809 "
        "J0810 J0811 J0812 J0813 J0814 J1607 J1608 J1609 J1610 J1611 J1612 J1613 J1614 J1615 J1616 J1617 "
        "J1618 J1619 J1620 J1621 J1622 J1623 J1624 J1625 J1626 J1627 J1628 J1629 J1630 J1631 J1632 J1633 "
        "J1634 J1635 J1636 J1637 J1638 J1639 J1640 J1641 J1642 J1643 J1644 J1645 J1646 J1647 J1648 J1649 "
        "J1650 J1651 J1652 J1653 J1654 J1655 J1656 J1657 J1658 J1659 J1660 J1661 J1662 J1663 J1664 J1665 "
        "J1666 J1667 J1668 J1669 J1670 J1671 J1689 J1682 J1680 J1683 J1684 J1681 J1687 J1688 J1686 J1685 "
        "J1672 J1673 J1674 J1705 J1706 J1707 J1708 J1675 J1676 J1677 J1678 J1679 J1690 J1691 J1692 J1693 "
        "J1694 J1695 J1696 J1697 J1698 J1699 L1700 J1701 J1702 J1703 J1704 J1709 J1718 J1717 J1710 J1711 "
        "J1712 J1713 J1714 J1715 J1716 J1755 J1762 J1725 J1731 J1751 J1723 J1730 J1733 J1752 J1735 J1738 "
        "J1740 J1743 J1744 J1749 J1756 J1757 J1764 J1763 J1761 J1760 J1759 J1758 J1754 J1753 J1750 J1748 "
        "J1747 J1746 J1745 J1742 J1741 J1739 J1737 J1736 J1734 J1732 J1726 J1724 J1722 J1721 J1720 J1719 "
        "J1729 J1728 J1727 J1766 J1767 J1768 J1765 J1769 J1770 J1772 J1771 J1773 J1774 J1775 J1776 J1777 "
        "J1778 J1779 J1780 J1781 J1782 J1783 J1784 J1785 J1786 J1787 J1788 J1789 J1790 J1791 J1792 J1793 "
        "J1794 J1795 J1796 J1797 J1798 J1799 J1800 J1801 J1802 J1803 J1804 J1805 J1806 J1807 J1808 J1809 "
        "J1810 J1811 J1812 J1837 J1836 J1814 J1838 J1815 J1816 J1817 J1818 J1819 J1820 J1821 J1822 J1823 "
        "J1824 J1825 J1826 J1827 J1828 J1829 J1830 J1813 J1831 J1832 J1833 J1834 J1835 J1842 J1843 J1844 "
        "J1845 J1846 J1847 J1848 J1849 J1850 J1851 J1852 J1853 J1839 J1840 J1841 J1854 J1855 J1856 J1857 "
        "J1859 J1860 J1861 J1862 J1865 J1869 J1870 J1871 J1876 J1877 J1878 J1917 J1884 J1882 J1881 J1880 "
        "J1942 J1866 J1867 J1868 J1864 J1863 J1858 J1872 J1873 J1874 J1875 J1879 J1883 J1885 J1886 J1887 "
        "J1888 J1889 J1890 J1891 J1892 J1893 J1894 J1895 J1896 J1897 J1898 J1900 J1901 J1902 J1903 J1904 "
        "J1905 J1906 J1907 J1908 J1909 J1912 J1913 J1914 J1915 J1916 J1932 J1933 J1934 J1935 J1936 J1937 "
        "J1938 J1939 J1940 J1941 J1993 J1994 J2003 J2002 J2012 J2027 J2065 J2080 J2016 J2024 J2025 J2026 "
        "J2035 J2036 J2040 J2046 J2051 J2053 J2058 J2063 J2064 J2066 J2067 J2068 J2069 J2074 J2081 J2082 "
        "J2083 J1919 J1920 J1922 J1923 J1899 J1910 J1911 J1918 J1921 J1924 J1925 J1926 J1927 J1928 J1929 "
        "J1930 J1931 J1943 J1944 J1945 J1946 J1947 J1948 J1949 J1950 J1951 J1952 J1953 J1954 J1955 J1956 "
        "J1957 J1958 J1959 J1960 J1961 J1962 J1963 J1964 J1965 J1966 J1967 J1968 J1969 J1970 J1971 J1972 "
        "J1973 J1974 J1975 J1976 J1977 J1978 J1979 J1980 J1981 J1982 J1983 J1984 J1985 J1986 J1987 J1988 "
        "J1989 J1990 J1991 J1992 J1995 J1996 J1997 J1998 J1999 J2000 J2001 J2060 J2084 J2004 J2005 J2006 "
        "J2007 J2008 J2009 J2010 J2011 J2013 J2014 J2015 J2017 J2018 J2019 J2020 J2021 J2022 J2023 J2028 "
        "J2029 J2030 J2031 J2032 J2033 J2034 J2037 J2038 J2039 J2041 J2042 J2043 J2044 J2045 J2047 J2048 "
        "J2049 J2050 J2052 J2054 J2055 J2056 J2057 J2059 J2061 J2062 J2070 J2071 J2072 J2073 J2075 J2076 "
        "J2077 J2078 J2079 J2114 J2113 J2112 J2111 J2110 J2109 J2108 J2107 J2106 J2105 J2104 J2103 J2102 "
        "J2101 J2100 J2099 J2098 J2097 J2096 J2095 J2094 J2093 J2092 J2091 J2090 J2089 J2088 J2087 J2086 "
        "J2085 K0001 K0013 K0008 K0010 K0030 K0020 K0018 K0032 K0031 K0022 K0025 K0005 K0024 K0017 K0007 "
        "K0011 K0009 K0002 K0021 K0033 K0035 K0034 K0023 K0041 K0042 K0039 K0036 K0038 K0040 K0050 K0049 "
        "K0046 K0048 K0047 K0043 K0044 K0045 K0003 K0004 K0006 K0015 K0012 K0014 K0016 K0019 K0028 K0029 "
        "K0026 K0027 K0037 K0051 K0058 K0070 K0071 K0072 K0073 K0074 K0075 K0076 K0079 K0077 K0078 K0069 "
        "K0068 K0054 K0056 K0066 K0067 K0083 K0084 K0085 K0086 K0087 K0063 K0053 K0052 K0055 K0064 K0065 "
        "K0062 K0061 K0088 K0089 K0090 K0091 K0092 K0059 K0060 K0057 K0080 K0081 K0082 K0093 K0094 K0095 "
        "K0096 K0097 K0098 K0101 K0102 K0103 K0104 K0105 K0106 K0107 K0108 K0109 K0110 K0111 K0112 K0113 "
        "K0114 K0115 K0116 K0117 K0118 K0119 K0120 K0121 K0122 K0123 K0124 K0125 K0126 K0127 K0128 K0129 "
        "K0130 K0099 K0100 S0006 S0005 K0238 K0239 K0240 K0241 K0242 K0243 K0244 K0245 K0246 K0247 K0248 "
        "K0249 K0250 K0251 K0252 K0253 K0254 K0255 K0256 K0257 K0258 K0234 K0231 K0232 K0233 K0235 K0236 "
        "K0237 K0267 K0266 K0259 K0260 K0261 K0262 K0263 K0264 K0265 K0286 K0285 K0284 K0287 K0283 K0282 "
        "K0279 K0277 K0274 K0271 K0269 K0299 K0296 K0298 K0300 K0301 K0302 K0303 K0305 K0307 K0308 K0273 "
        "K0297 K0288 K0289 K0268 K0272 K0278 K0281 K0291 K0292 K0293 K0290 K0276 K0275 K0280 K0270 K0294 "
        "K0295 K0304 K0310 K0306 K0309 K0341 K0340 K0339 K0337 K0333 K0332 K0338 K0335 K0334 K0336 K0326 "
        "K0325 K0323 K0324 K0317 K0315 K0316 K0313 K0314 K0311 K0312 K0322 K0321 K0320 K0319 K0318 K0328 "
        "K0327 K0331 K0329 K0330 K0342 S0010 S0009 "
    ),
    63: (
        "K0132 K0131 K0133 K0134 K0135 K0136 K0137 K0138 K0139 K0140 K0141 K0142 K0143 K0144 K0145 K0146 "
        "K0147 K0148 K0149 K0150 K0151 K0152 K0153 K0154 K0155 K0157 K0156 K0163 K0159 K0160 K0162 K0161 "
        "K0164 K0167 K0165 K0173 K0168 K0169 K0170 K0171 K0172 K0176 K0158 K0379 K0175 K0166 K0178 K0177 "
        "K0179 K0180 K0181 K0182 K0183 K0184 K0185 K0186 K0187 K0188 K0189 K0190 K0191 K0192 K0193 K0194 "
        "K0195 K0196 K0197 K0198 K0199 K0200 K0201 K0202 K0203 K0204 K0206 K0205 K0207 K0210 K0208 K0209 "
        "K0211 K0212 K0213 K0214 K0215 K0443 K0444 K0445 K0446 K0447 K0448 K0449 K0450 K0451 K0452 K0453 "
        "K0454 K0455 K0216 K0218 K0219 K0220 K0221 K0222 K0456 K0457 K0458 K0459 K0465 K0344 K0343 K0225 "
        "K0217 K0226 K0224 K0460 K0463 K0464 K0461 K0462 K0229 K0228 K0230 K0227 K0223 K0389 K0386 K0390 "
        "K0392 K0394 K0391 K0368 K0361 K0360 K0359 K0358 K0357 K0356 K0355 K0353 K0351 K0352 K0350 K0348 "
        "K0347 K0346 K0345 K0367 K0372 K0365 K0362 K0375 K0473 K0469 K0471 K0466 K0468 K0467 K0470 K0475 "
        "K0476 K0474 K0472 K0385 K0378 K0387 K0380 K0381 K0382 K0383 K0384 K0484 K0479 K0481 K0480 K0485 "
        "K0483 K0482 K0477 K0478 K0487 K0488 K0397 K0377 K0369 K0374 K0366 K0373 K0364 K0363 K0370 K0376 "
        "K0388 K0393 K0395 K0396 K0354 K0349 K0371 K0489 K0491 K0490 K0398 K0486 K0399 K0400 K0401 K0402 "
        "K0403 K0404 K0405 K0174 K0492 K0406 K0412 K0409 K0408 K0407 K0411 K0410 K0414 K0416 K0420 K0421 "
        "K0415 K0417 K0413 K0422 K0419 K0418 K0435 K0436 K0438 K0439 K0440 K0441 K0442 K0431 K0432 K0433 "
        "K0434 K0430 K0429 K0427 K0426 K0424 K0428 K0425 K0630 K0631 K0423 K0648 K0647 K0646 K0645 K0644 "
        "K0643 K0642 K0641 K0640 K0632 K0633 K0634 K0493 K0494 K0496 K0495 K0497 K0498 K0500 K0502 K0501 "
        "K0499 K0437 K0649 K0655 K0656 K0653 K0651 K0650 K0652 K0654 K0667 K0668 K0636 K0637 K0639 K0635 "
        "K0638 K0666 K0503 K0505 K0506 K0504 K0660 K0658 K0663 K0664 K0662 K0657 K0661 K0665 K0659 K0507 "
        "K0509 K0508 K0510 K0511 K0512 K0513 K0514 K0515 K0675 K0674 K0673 K0672 K0671 K0670 K0669 K0516 "
        "K0686 K0684 K0685 K0683 K0682 K0676 K0677 K0678 K0680 K0679 K0681 K0519 K0517 K0520 K0521 K0518 "
        "K0694 K0696 K0693 K0695 K0529 K0527 K0528 K0524 K0523 K0522 K0525 K0526 K0688 K0691 K0687 K0690 "
        "K0692 K0689 K0535 K0534 K0533 K0530 K0531 K0532 K0699 K0697 K0698 K0536 K0700 K0702 K0701 K0703 "
        "K0537 K0540 K0541 K0707 K0709 K0705 K0710 K0708 K0706 K0704 K0542 K0539 K0538 K0544 K0711 K0713 "
        "K0712 K0546 K0545 K0547 K0543 K0549 K0548 K0552 K0550 K0551 K0553 K0714 K0715 K0716 K0717 K0718 "
        "K0719 K0720 K0554 K0725 K0721 K0722 K0724 K0723 K0555 K0556 K0557 K0595 K0599 K0596 K0597 K0593 "
        "K0594 K0598 K0600 K0559 K0560 K0558 K0602 K0603 K0604 K0601 K0605 K0606 K0607 K0608 K0609 K0619 "
        "K0618 K0617 K0616 K0615 K0614 K0613 K0612 K0620 K0611 K0610 K0563 K0568 K0567 K0569 K0562 K0564 "
        "K0565 K0561 K0566 K0621 K0622 K0623 K0624 K0625 K0572 K0573 K0574 K0570 K0571 K0577 K0575 K0576 "
        "K0729 K0727 K0730 K0731 K0728 K0627 K0628 K0629 K0626 K0726 K0585 K0586 K0583 K0584 K0582 K0581 "
        "K0580 K0579 K0578 K0589 K0588 K0587 K0591 K0590 K0592 K0737 K0738 K0735 K0732 K0734 K0733 K0739 "
        "K0742 K0741 K0740 K0736 K0743 K0744 "
    ),
    64: (
        "J0910 J0911 J0912 J0913 J0914 J0915 J0916 J0917 J0918 J0919 J0920 J0921 J0922 J0923 J0924 J0925 "
        "J0926 J0927 J0928 J0929 J0930 J0931 J0932 J0933 J0934 J0935 J0936 J0937 J0938 J0939 J0940 J0941 "
        "J0942 J0943 J0944 J0945 J0946 J0947 J0948 J0949 J0950 J0951 J0952 J0953 J0954 J0955 J0956 J0957 "
        "J0958 J0959 J0960 J0961 J0962 J0994 J0995 J0996 J0997 J0963 J0964 J0965 J0966 J0967 J0968 J0969 "
        "J0970 J0971 J0973 J0975 J0977 J0998 J0999 J1000 J1001 J1002 J0972 J0974 J0976 J0978 J0979 J1004 "
        "J0981 J0982 J0983 J0980 J1003 J1006 J0986 J0985 J0984 J1005 J1030 J1031 J1032 J0987 J1007 J1054 "
        "J1057 J1052 J1053 J1055 J1056 J1033 J1034 J1035 J1036 J1037 J1038 J1039 J1040 J1041 J1042 J1043 "
        "J1044 J1045 J1046 J1047 J1048 J1051 J1133 J1132 J1131 J1129 J1058 J1059 J1060 J1061 J1062 J1063 "
        "J1064 J1065 J1066 J1067 J1068 J1069 J1070 J1072 J1071 J1207 J1206 J1204 J1203 J1202 J1200 J1199 "
        "J1198 J1197 J1196 J1145 J1127 J1128 J1130 J1134 J1135 J1137 J1138 J1139 J1140 J1141 J1144 J1136 "
        "J1201 J1142 J1146 J1147 J1148 J1149 J1150 J1008 J1073 J1074 J1075 J1076 J1077 J1078 J1079 J1208 "
        "J1209 J1210 J1211 J1212 J1213 J1080 J1081 J1082 J1083 J1084 J1143 J1151 J1152 J1153 J1154 J1155 "
        "J1156 J1157 J1205 J1214 J1215 J1216 J1217 J1218 J1219 J1220 J1223 J0989 J0990 J1009 J1010 J1167 "
        "J1221 J1162 J1163 J1160 J1222 J1164 J1228 J0988 J1226 J1085 J1225 J1161 J1159 J1165 J1224 J1166 "
        "J1230 J1227 J1229 J1088 J1086 J1087 J1090 J1089 J1091 J1231 J1234 J1233 J1168 J1232 J1158 J1169 "
        "J1095 J1099 J1100 J1096 J1094 J1092 J1093 J1238 J1236 J1098 J1239 J1240 J1097 J1237 J1235 J1170 "
        "J1171 J1101 J1102 J1103 J1172 J1173 J1174 J1175 J1176 J1177 J1178 J1179 J1180 J1181 J1241 J1242 "
        "J1011 J1104 J1105 J1106 J1107 J1108 J1182 J1183 J1184 J1185 J1243 J1244 J1109 J1110 J1111 J1112 "
        "J1113 J1114 J1115 J1116 J1117 J1118 J1119 J1120 J1121 J1122 J1186 J1187 J1188 J1189 J1190 J1191 "
        "J1245 J1246 J1247 J1248 J1249 J1192 J1193 J1194 J1195 J1250 J1251 J1252 J1253 J1254 J1257 J1256 "
        "J1258 J1259 J1260 J1261 J1262 J1265 J1266 J1123 J1124 J1125 J1255 J1263 J1264 J1268 J1126 J1271 "
        "J1269 J1270 J1267 J1280 J1281 J1282 J1283 J1284 J1285 J1286 J1287 J1288 J1289 J1290 J1291 J1292 "
        "J1293 J1294 J1303 J1304 J1305 J1306 J1307 J1327 J1328 J1329 J1330 J1331 J1332 J1333 J1334 J1335 "
        "J1412 J1413 J1414 J1272 J1273 J1274 J1275 J1276 J1277 J1278 J1279 J1302 J1298 J1299 J1300 J1301 "
        "J1297 J1296 J1295 J0991 J0992 J1012 J1013 J1014 J1015 J1016 J1017 J1018 J1019 J1020 J1021 J1022 "
        "J1023 J1024 J1025 J1026 J1415 J1416 J1417 J1418 J1419 J1420 J1421 J1422 J1423 J1424 J1425 J1308 "
        "J1309 J1310 J1311 J1312 J1313 J1314 J1315 J1316 J1317 J1318 J1319 J1336 J1337 J1338 J1339 J1340 "
        "J1341 J1342 J1343 J1428 J1427 J1348 J1347 J1346 J1345 J1344 J1429 J1430 J1431 J1432 J1433 J1434 "
        "J1435 J1436 J1027 J0993 J1426 J1320 J1437 J1438 J1439 J1440 J1441 J1442 J1443 J1321 J1444 J1445 "
        "J1446 J1447 J1028 J1322 J1029 J1323 J1324 J1349 J1350 J1351 J1352 J1353 J1354 J1355 J1356 J1357 "
        "J1358 J1448 J1449 J1450 J1451 J1452 J1453 J1454 J1455 J1456 J1457 J1458 J1459 J1460 J1461 J1462 "
        "J1463 J1464 J1465 J1466 J1467 J1468 J1469 J1470 J1471 J1472 J1473 J1409 J1410 J1474 J1475 J1476 "
        "J1477 J1478 J1359 J1360 J1361 J1411 J1362 J1363 J1364 J1479 J1480 J1481 J1482 J1483 J1484 J1485 "
        "J1486 J1487 J1488 J1489 J1490 J1491 J1492 J1493 J1494 J1495 J1496 J1497 J1498 J1499 J1500 J1501 "
        "J1502 J1503 J1504 J1365 J1366 J1367 J1368 J1369 J1370 J1371 J1372 J1373 J1374 J1505 J1506 J1507 "
        "J1508 J1509 J1510 J1511 J1512 J1513 J1514 J1515 J1516 J1517 J1518 J1519 J1520 J1521 J1522 J1523 "
        "J1524 J1525 J1526 J1527 J1528 J1529 J1530 J1531 J1325 J1326 J1375 J1376 J1377 J1378 J1379 J1380 "
        "J1381 J1382 J1383 J1384 J1385 J1386 J1387 J1388 J1389 J1390 J1391 J1392 J1393 J1394 J1395 J1396 "
        "J1397 J1398 J1532 J1533 J1534 J1535 J1536 J1537 J1538 J1539 J1540 J1541 J1543 J1542 J1544 J1545 "
        "J1546 J1399 J1400 J1401 J1402 J1403 J1404 J1405 J1406 J1407 J1547 J1548 J1549 J1550 J1551 J1552 "
        "J1408 J1553 J1554 J1555 J1556 J1557 J1558 J1559 J1560 J1561 J1562 J1563 J1564 J1565 J1566 J1567 "
        "J1584 J1585 J1586 J1568 J1569 J1570 J1571 J1572 J1573 J1574 J1575 J1576 J1577 J1578 J1579 J1580 "
        "J1581 J1583 J1582 "
    ),
    67: (
        "J0051 J0052 J0053 J0054 J0055 J0056 J0057 J0058 "
    ),
    69: (
        "J0501 J0502 J0503 J0504 J0505 J0506 J0507 J0508 J0509 J0510 J0511 J0512 J0513 J0515 J0516 J0517 "
        "J0518 J0519 J0520 J0521 J0523 J0525 J0526 J0527 J0528 J0529 J0530 J0531 J0532 J0535 J0536 J0541 "
        "J0545 J0548 J0550 J0551 J0552 J0553 J0554 J0555 J0556 J0557 J0558 J0559 J0560 J0561 J0562 J0563 "
        "J0564 J0565 J0566 J0567 J0568 J0569 J0571 J0572 J0573 J0574 J0575 J0576 J0577 J0578 J0579 J0580 "
        "J0581 J0582 J0583 J0584 J0585 J0586 J0587 J0588 J0589 J0590 J0591 J0592 J0593 J0594 J0595 J0596 "
        "J0597 J0598 J0599 J0710 J0711 J0712 J0713 J0714 J0715 J0716 J0717 J0718 J0719 J0720 J0721 J0722 "
        "J0723 J0724 J0725 J0726 J0727 J0728 J0729 J0730 J0731 J0732 J0733 J0734 J0735 J0736 J0737 J0738 "
        "J0739 J0740 J0741 J0742 J0743 J0744 J0745 J0746 J0747 J0748 J0749 J0750 J0752 J0753 J0754 J0755 "
        "J0756 J0757 J0758 J0759 J0772 J0779 J0780 J0781 J0798 J0799 J0802 J0815 J0816 J0817 J0818 J0819 "
        "J0820 J0821 J0822 J0823 J0824 J0825 J0826 J0830 J0831 J0832 J0833 J0834 J0835 J0836 J0837 J0838 "
        "J0839 J0840 J0841 J0842 J0843 J0844 J0845 J0846 J0848 J0849 J0850 J0851 J0852 J0853 J0854 J0855 "
        "J0856 J0857 J0859 J0860 J0861 J0862 J0863 J0864 J0865 J0866 J0867 J0868 J0869 J0870 J0871 J0872 "
        "J0873 J0874 J0875 J0876 J0877 J0878 J0879 J0880 J0881 J0882 J0883 J0884 J0885 J0886 J0887 J0888 "
        "J0889 J0890 J0891 J0892 J0893 J0894 J0895 J0896 J0897 J0898 J0899 J0900 J0901 J0902 J0903 J0904 "
        "J0905 J0906 J0907 J0908 J0909 "
    ),
    70: (
        "J0351 J0352 J0353 J0354 J0355 J0356 J0357 J0358 J0359 J0360 J0361 J0362 J0363 J0364 J0365 J0366 "
        "J0367 J0368 J0369 J0370 J0371 J0372 J0373 J0374 J0375 J0376 J0377 J0378 J0379 J0380 J0381 J0382 "
        "J0383 J0384 J0385 J0401 J0402 J0403 J0404 J0405 J0406 J0407 J0408 J0409 J0410 J0411 J0412 J0413 "
        "J0414 J0415 J0416 J0417 J0418 J0419 J0420 J0421 J0422 J0423 J0424 J0425 J0426 J0427 J0428 J0429 "
        "J0430 J0431 J0432 J0433 J0434 J0435 J0436 J0437 J0438 J0439 J0440 J0441 J0442 J0443 J0444 J0445 "
        "J0446 J0447 J0448 J0449 J0450 J0451 J0452 J0453 J0454 J0455 J0456 J0458 J0462 J0463 J0465 J0466 "
        "J0470 J0471 J0474 J0475 J0476 J0477 J0478 J0479 J0483 J0485 J0492 J0493 J0494 J0495 J0496 J0497 "
    ),
    71: (
        "J0300 "
    ),
    72: (
        "J0386 J0387 J0388 J0389 J0390 J0391 J0392 J0393 J0394 J0395 J0396 J0397 J0398 J0399 J0400 "
    ),
    73: (
        "I0001 I0451 I0504 I0508 I0510 I0511 I0515 I0536 I0538 I0547 I0602 I0706 I0709 I0710 I0712 I0714 "
        "I0716 I0717 I0718 I0720 I0721 I0722 I0723 I0726 I0730 I0732 I0735 I0740 I0744 I0745 I0747 I0748 "
        "I0757 I0759 I0766 I0767 I0771 I0772 I0773 I0775 I0776 I0779 I0782 I0784 I0792 I0795 I0796 I0903 "
        "I0905 I0909 I0910 I0913 I0915 I0920 I0921 I0926 I0927 I0934 I0935 I0936 I0937 I0938 I0943 I0951 "
        "I0952 I0955 I0957 I0960 I0968 I0970 I0975 I0985 I0986 I0989 I0992 I0994 I0996 I0997 I0998 I1001 "
        "I1002 I1006 I1007 I1011 I1014 I1020 I1022 I1025 I1028 I1029 I1033 I1035 I1040 I1042 I1043 I1044 "
        "I1045 I1047 I1049 I1050 I1053 I1055 I1056 I1060 I1062 I1064 I1065 I1066 I1067 I1069 I1073 I1075 "
        "I1078 I1079 I1081 I1086 I1089 I1091 I1096 I1097 I1102 I1103 I1104 I1109 I1111 I1112 I1113 I1114 "
        "I1117 I1118 I1120 I1125 I1126 I1128 I1130 I1134 I1135 I1139 I1140 I1141 I1143 I1145 I1147 I1148 "
        "I1149 I1151 I1152 I1154 I1155 I1156 I1158 I1161 I1164 I1173 I1176 I1177 I1178 I1181 I1186 I1187 "
        "I1190 I1192 I1277 I1279 I1280 I1281 I1285 I1292 I1293 I1294 I1298 I1299 I1302 I1310 I1311 I1312 "
        "I1314 I1316 I1319 I1320 I1321 I1322 I1323 I1324 I1325 I1326 I1328 I1336 I1337 I1338 I1342 I1346 "
        "I1347 I1348 I1349 I1350 I1352 I1354 I1357 I1360 I1361 I1362 I1368 I1373 I1379 I1381 I1383 I1392 "
        "I1393 I1395 I1398 I1400 I1406 I1413 I1416 I1420 I1430 I1431 I1432 I1433 I1435 I1442 I1445 I1446 "
        "I1448 I1449 I1451 I1452 I1453 I1454 I1455 I1456 I1458 I1460 I1461 I1462 I1463 I1464 I1465 I1466 "
        "I1469 I1470 I1474 I1475 I1601 I1602 I1603 I1604 I1605 I1607 I1608 I1609 I1611 I1612 I1613 I1614 "
        "I1615 I1616 I1618 I1621 I1622 I1623 I1624 I1625 I1626 I1627 I1633 I1636 I1638 I1639 I1640 I1644 "
        "I1645 I1646 I1647 I1651 I1652 I1653 I1654 I1658 I1659 I1660 I1661 I1665 I1666 I1667 I1668 I1669 "
        "I1671 I1672 I1673 I1674 I1675 I1676 I1677 I1678 I1680 I1681 I1682 I1683 I1684 I1685 I1686 I1688 "
        "I1689 I1690 I1691 I1692 I1693 I1694 I1696 I1697 I1698 I1699 I1700 I1901 I1902 I1903 I1905 I1906 "
        "I1907 I1908 I1909 I1910 I1911 I1912 I1913 I1914 I1917 I1919 I1921 I1922 I1923 I1926 I1927 I1928 "
        "I1935 I1938 I1939 I1941 I1942 I1944 I1945 I1947 I1948 I1949 I1950 I1951 I1953 I1954 I1955 I1957 "
        "I1958 I1960 I1961 I1963 I1965 I1966 I1967 I1970 I1971 I1972 I1979 I1980 I1981 I1983 I1984 I1987 "
        "I1990 I1995 I1996 I1998 I1999 I2000 I2021 I2022 I2025 I2026 I2027 I2028 I2029 I2033 I2034 I2035 "
        "I2039 I2042 I2043 I2044 I2045 I2046 I2047 I2048 I2049 I2050 I2053 I2054 I2055 I2057 I2061 I2062 "
        "I2063 I2066 I2069 I2070 I2071 I2072 I2073 I2901 I2902 I2903 I2904 I2905 I2906 I2907 I2908 I2909 "
        "I2910 I2911 I2912 I2915 I2917 I2919 I2921 I2922 I2923 I2924 I2925 I2926 I2927 I2928 I2930 I2931 "
        "I2932 I2935 I2936 I2937 I2939 I2942 I2943 I2944 I2945 I2946 I2947 I2948 I2949 I2950 I2952 I2953 "
        "I2954 I2955 I2956 I2958 I2959 I2960 I2961 I2962 I2963 I2964 I2965 I2966 I2968 I2970 I2972 I2973 "
        "I2974 I2975 I2976 I2977 I2978 I2979 I2980 I2989 I2992 I2993 I2994 I2995 I2996 I2997 I2998 I2999 "
        "I3000 I3309 I3310 I3311 I3312 I3313 I3314 I3315 I3316 I3317 I3318 I3319 I3320 I3321 I3322 I3323 "
        "I3324 I3325 I3326 I3327 I3328 I3329 I3330 I3331 I3332 I3333 I3334 I3335 I3336 I3337 I3338 I3339 "
        "I3340 I3341 I3342 I3343 I3344 I3345 I3346 I3347 I3348 I3349 I3350 I3351 I3352 I3353 I3354 I3355 "
        "I3356 I3357 I3358 I3359 I3360 I3361 I3362 I3363 I3364 I3365 I3366 I3367 I3368 I3369 I3370 I3371 "
        "I3372 I3373 I3374 I3375 I3376 I3377 I3378 I3379 I3380 I3381 I3382 I3383 I3384 I3385 I3386 I3387 "
        "I3388 I3389 I3390 I3391 I3392 I3393 I3394 I3395 I3396 I3397 I3398 I3399 I3400 J0001 J0002 J0003 "
        "J0004 J0005 J0006 J0007 J0008 J0009 J0010 J0011 J0012 J0013 J0014 J0015 J0016 J0017 J0018 J0019 "
        "J0020 J0021 J0022 J0023 J0024 J0025 J0026 J0027 J0028 J0029 J0030 J0031 J0032 J0033 J0034 J0035 "
        "J0036 J0037 J0038 J0039 J0040 J0041 J0042 J0043 J0044 J0045 J0046 J0047 J0048 J0049 J0050 J0059 "
        "J0060 J0061 J0062 J0063 J0064 J0065 J0066 J0067 J0068 J0069 J0070 J0071 J0072 J0073 J0074 J0075 "
        "J0076 J0077 J0078 J0079 J0080 J0081 J0082 J0083 J0084 J0085 J0086 J0087 J0088 J0089 J0090 J0091 "
        "J0092 J0093 J0094 J0095 J0096 J0097 J0098 J0099 J0100 J0101 J0102 J0103 J0105 J0106 J0107 J0108 "
        "J0109 J0110 J0111 J0112 J0113 J0114 J0115 J0116 J0117 J0118 J0119 J0120 J0121 J0122 J0123 J0124 "
        "J0125 J0126 J0127 J0128 J0129 J0130 J0131 J0132 J0133 J0134 J0135 J0136 J0137 J0140 J0141 J0142 "
        "J0143 J0145 J0146 J0147 J0148 J0149 J0150 J0151 J0152 J0153 J0154 J0155 J0156 J0157 J0158 J0159 "
        "J0160 J0162 J0163 J0164 J0166 J0167 J0168 J0169 J0170 J0171 J0172 J0173 J0174 J0175 J0176 J0177 "
        "J0178 J0179 J0180 J0181 J0182 J0183 J0184 J0185 J0186 J0187 J0188 J0189 J0190 J0191 J0192 J0193 "
        "J0194 J0195 J0196 J0197 J0198 J0199 J0200 J0201 J0202 J0203 J0204 J0205 J0206 J0207 J0208 J0209 "
        "J0210 J0211 J0212 J0213 J0214 J0215 J0216 J0217 J0218 J0219 J0220 J0221 J0222 J0223 J0224 J0225 "
        "J0226 J0227 J0228 J0229 J0230 J0231 J0232 J0233 J0234 J0235 J0236 J0237 J0238 J0239 J0240 J0241 "
        "J0242 J0243 J0244 J0245 J0246 J0247 J0248 J0249 I1191 I2056 I1920 J0104 J0161 J0165 J0138 J0139 "
    ),
}

SAMPLE_FILE_PATH = "sample_etm_data.csv"


@st.cache_data(show_spinner=False)
def get_embedded_fleet_lookup():
    """Builds the embedded MTC fleet capacity lookup DataFrame (3,390 vehicles)."""
    rows = []
    for capacity, fleet_str in FLEET_CAPACITY_GROUPS.items():
        for fleet_no in fleet_str.split():
            rows.append((fleet_no, capacity))
    return pd.DataFrame(rows, columns=["fleet_no", "total_capacity"])


# Custom CSS for styling
st.markdown("""
    <style>
    .stApp {
        background: linear-gradient(135deg, #f5f7fa 0%, #e4edf5 100%);
    }
    .header {
        color: #2c3e50;
        text-align: center;
        padding: 1rem;
        font-size: 2.5rem;
        font-weight: 700;
        margin-bottom: 0.5rem;
        background: white;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .subscript {
        text-align: center;
        color: #5d6d7e;
        font-weight: 600;
        margin-bottom: 1.5rem;
    }
    </style>
""", unsafe_allow_html=True)

# Initialize session state
if 'analysis_complete' not in st.session_state:
    st.session_state.analysis_complete = False
if 'all_service' not in st.session_state:
    st.session_state.all_service = pd.DataFrame()
if 'all_route' not in st.session_state:
    st.session_state.all_route = pd.DataFrame()
if 'all_sample' not in st.session_state:
    st.session_state.all_sample = pd.DataFrame()
if 'selected_day_type' not in st.session_state:
    st.session_state.selected_day_type = ""

# Header
st.markdown('<div class="header">ETM Data Analysis Platform PS</div>', unsafe_allow_html=True)
st.markdown('<div class="subscript">By Sarath Kumar Ponugumati</div>', unsafe_allow_html=True)

# Load embedded fleet lookup
lookup_df = get_embedded_fleet_lookup()


def parse_date(date_str):
    try:
        date_str = str(date_str).strip()
        for fmt in ('%Y-%m-%d', '%d-%m-%Y', '%d/%m/%Y', '%m/%d/%Y', '%d-%m-%y'):
            try:
                return datetime.strptime(date_str, fmt)
            except ValueError:
                continue
        return pd.to_datetime(date_str, dayfirst=True)
    except Exception:
        return None


def read_etm_source(source):
    """Reads an ETM file from either a file path string or a Streamlit UploadedFile."""
    if isinstance(source, str):
        return pd.read_csv(source, compression='infer')
    source.seek(0)
    return pd.read_csv(source, compression='infer')


def extract_file_date(file_name, source):
    """Extracts the date from the filename first, falling back to reading the first row inside the file."""
    try:
        if "Ticket-Data" in file_name:
            date_part = file_name.split('-')[2:5]
            date_str = '-'.join(date_part).split('.')[0]
            return datetime.strptime(date_str, '%Y-%m-%d').date()
        elif '_' in file_name:
            date_str = file_name.split('_')[1].split('.')[0]
            if len(date_str) == 8 and date_str.isdigit():
                return datetime.strptime(date_str, '%d%m%Y').date()
    except Exception:
        pass

    try:
        if isinstance(source, str):
            sample_df = pd.read_csv(source, compression='infer', nrows=1)
        else:
            source.seek(0)
            sample_df = pd.read_csv(source, compression='infer', nrows=1)
            source.seek(0)
        date_col = 'date' if 'date' in sample_df.columns else ('duty_date' if 'duty_date' in sample_df.columns else None)
        if date_col and not sample_df.empty:
            parsed = parse_date(sample_df[date_col].iloc[0])
            if parsed is not None and not pd.isna(parsed):
                return parsed.date()
    except Exception:
        pass
    return None


# Sidebar Configuration
with st.sidebar:
    st.header("⚙️ Data Source & Info")
    data_source = st.radio(
        "Select Data Source",
        ["Use Sample ETM Data", "Upload ETM File(s)"],
        index=0
    )
    st.caption(f"✅ Fleet lookup built-in ({len(lookup_df):,} vehicles; missing fleet capacity = 65)")

# Collect available files as list of (display_name, file_source)
available_items = []
if data_source == "Use Sample ETM Data":
    if os.path.exists(SAMPLE_FILE_PATH):
        available_items = [(SAMPLE_FILE_PATH, SAMPLE_FILE_PATH)]
        st.info(f"Loaded built-in sample ETM dataset: `{SAMPLE_FILE_PATH}`")
    else:
        st.error(f"Sample file `{SAMPLE_FILE_PATH}` was not found in the repository root. Please upload `{SAMPLE_FILE_PATH}` to GitHub or switch to 'Upload ETM File(s)'.")
else:
    uploaded_files = st.file_uploader(
        "Upload one or more ETM files (.csv or .bz2)",
        type=["csv", "bz2"],
        accept_multiple_files=True
    )
    if uploaded_files:
        available_items = [(uf.name, uf) for uf in uploaded_files]
        st.info(f"Uploaded {len(available_items)} ETM file(s)")

# Main Content: Analysis Method
st.header("🔍 Select Analysis Method")
analysis_method = st.radio(
    "Choose how to filter files:",
    [
        "Process all selected files",
        "Process files by date range",
        "Process files by month",
        "Process files by weekday/weekend"
    ],
    index=0
)

# Map files to dates when date-based filtering is selected
item_date_map = {}
if available_items and analysis_method != "Process all selected files":
    for name, src in available_items:
        dt = extract_file_date(name, src)
        if dt is not None:
            item_date_map[name] = (dt, src)

# Date range selection
filtered_items = []
if analysis_method == "Process files by date range" and available_items:
    st.subheader("📅 Select Date Range")
    if item_date_map:
        dates = [v[0] for v in item_date_map.values()]
        col1, col2 = st.columns(2)
        with col1:
            start_date = st.date_input("Start Date", min(dates))
        with col2:
            end_date = st.date_input("End Date", max(dates))
        filtered_items = [(name, src) for name, (dt, src) in item_date_map.items() if start_date <= dt <= end_date]
        st.info(f"Found {len(filtered_items)} file(s) in selected date range")
    else:
        st.warning("No valid dates found inside the files.")

# Monthly analysis selection
elif analysis_method == "Process files by month" and available_items:
    st.subheader("🗓️ Select Month")
    if item_date_map:
        month_year_options = {}
        for name, (dt, src) in item_date_map.items():
            month_year = f"{calendar.month_name[dt.month]} {dt.year}"
            month_year_options.setdefault(month_year, []).append((name, src))
        sorted_months = sorted(month_year_options.keys(), key=lambda x: datetime.strptime(x, '%B %Y'))
        selected_month = st.selectbox("Available Months", sorted_months)
        filtered_items = month_year_options.get(selected_month, [])
        st.info(f"Found {len(filtered_items)} file(s) for {selected_month}")
    else:
        st.warning("No valid monthly data found inside the files.")

# Weekday/Weekend analysis selection
elif analysis_method == "Process files by weekday/weekend" and available_items:
    st.subheader("📅 Select Day Type")
    weekday_items = [(name, src) for name, (dt, src) in item_date_map.items() if dt.weekday() < 5]
    weekend_items = [(name, src) for name, (dt, src) in item_date_map.items() if dt.weekday() >= 5]

    st.info(f"Found {len(weekday_items)} weekday file(s) and {len(weekend_items)} weekend file(s)")

    col1, col2 = st.columns(2)
    with col1:
        if st.button("Weekdays (Mon-Fri)", key="weekdays", type="primary" if st.session_state.selected_day_type == "weekdays" else "secondary"):
            st.session_state.selected_day_type = "weekdays"
    with col2:
        if st.button("Weekends (Sat-Sun)", key="weekends", type="primary" if st.session_state.selected_day_type == "weekends" else "secondary"):
            st.session_state.selected_day_type = "weekends"

    if st.session_state.selected_day_type == "weekdays":
        filtered_items = weekday_items
        st.success(f"Selected Weekdays: {len(filtered_items)} file(s)")
    elif st.session_state.selected_day_type == "weekends":
        filtered_items = weekend_items
        st.success(f"Selected Weekends: {len(filtered_items)} file(s)")


# Main processing function
def process_file(file_name, file_source, lookup_df):
    try:
        try:
            df = read_etm_source(file_source)
        except Exception as e:
            st.error(f"Error reading {file_name}: {str(e)}")
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

        required_cols = {'schedule_no', 'fleet_no', 'from_stage', 'to_stage', 'ticket_type'}
        if not required_cols.issubset(df.columns):
            st.error(f"Missing required columns in {file_name}")
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

        # Process schedule numbers
        try:
            split_result = df['schedule_no'].astype(str).str.split('-', expand=True)
            if split_result.shape[1] >= 4:
                df[['service_type', 'route_no', 'service_id', 'shift']] = split_result.iloc[:, :4]
                df['service_no'] = df['route_no'].astype(str) + "-" + df['service_id'].astype(str)
                df['trip_shift'] = df['trip_no'].astype(str) + "-" + df['shift'].astype(str)
                df['service_trip_shift'] = df['service_no'].astype(str) + "_" + df['trip_shift'].astype(str) + "_" + df['service_type'].astype(str)
            else:
                st.error(f"Invalid schedule_no format in {file_name}")
                return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
        except Exception as e:
            st.error(f"Error processing schedule numbers in {file_name}: {str(e)}")
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

        # Convert stage columns to numeric
        df['to_stage'] = pd.to_numeric(df['to_stage'], errors='coerce')
        df['from_stage'] = pd.to_numeric(df['from_stage'], errors='coerce')

        # Removing dummy trips with KM zero
        df['KM'] = pd.to_numeric(df['KM'], errors='coerce')
        df = df[df['KM'].ne(0) | df['KM'].isna()].copy()

        # Total Stages per each route
        df['route_description'] = df['Trip_Origin'].astype(str) + " To " + df['Trip_Destination'].astype(str)
        df['route_desc'] = df['route_no'].astype(str) + "-" + df['route_description'].astype(str)

        df['max_stage'] = df[['to_stage', 'from_stage']].max(axis=1)
        df['total_stages'] = df.groupby('route_desc')['max_stage'].transform('max')
        df.drop(columns=['max_stage'], inplace=True)

        # Parse dates
        base_date = parse_date(df["date"].iloc[0])
        if base_date is not None and not pd.isna(base_date):
            date_str = base_date.strftime('%d-%m-%y')
            day_str = base_date.strftime('%A')
        else:
            date_str = str(df["date"].iloc[0])
            day_str = "Unknown"

        # Sample Dataframe creation with Ticket Counts
        ticket_counts = {
            'date': date_str,
            'day': day_str,
            'TC01_Adult': len(df[df['ticket_type'] == 'Adult Ticket']),
            'TC02_Difference': len(df[df['ticket_type'] == 'Difference Fare Ticket']),
            'TC03_Dis_esc': len(df[df["ticket_type"] == "Disable Person Escort Ticket"]),
            'TC04_Dis': len(df[df["ticket_type"] == "Disable Person Ticket"]),
            'TC05_Luggage': len(df[df["ticket_type"] == "Luggage Ticket"]),
            'TC06_Senior': len(df[df["ticket_type"] == "Senior Citizen Ticket"]),
            'TC07_Transgender': len(df[df['ticket_type'] == 'Transgender Ticket']),
            'TC08_Women': len(df[df["ticket_type"] == "Women Ticket"]),
            'all_tickets': len(df)
        }
        sample_size = pd.DataFrame([ticket_counts])

        # Passenger Counts for Each type of Ticket
        passenger_counts = df.groupby('ticket_type')['passenger_count'].sum()
        sample_size['PC01'] = passenger_counts.get('Adult Ticket', 0)
        sample_size['PC02'] = passenger_counts.get('Difference Fare Ticket', 0)
        sample_size['PC03'] = passenger_counts.get('Disable Person Escort Ticket', 0)
        sample_size['PC04'] = passenger_counts.get('Disable Person Ticket', 0)
        sample_size['PC06'] = passenger_counts.get('Senior Citizen Ticket', 0)
        sample_size['PC07'] = passenger_counts.get('Transgender Ticket', 0)
        sample_size['PC08'] = passenger_counts.get('Women Ticket', 0)
        sample_size['total_passengers'] = df['passenger_count'].sum()

        # Revenue Calculations
        revenue_counts = df.groupby('ticket_type')['total_ticket_amount'].sum()
        luggage_revenue = df[df['ticket_type'] == 'Luggage Ticket']['luggage_amount'].sum()

        sample_size['R01'] = revenue_counts.get('Adult Ticket', 0)
        sample_size['R02'] = revenue_counts.get('Difference Fare Ticket', 0)
        sample_size['R03'] = sample_size['PC03'] * 16
        sample_size['R04'] = sample_size['PC04'] * 16
        sample_size['R05'] = luggage_revenue
        sample_size['R06'] = sample_size['PC06'] * 21
        sample_size['R07'] = sample_size['PC07'] * 16
        sample_size['R08'] = sample_size['PC08'] * 16
        sample_size['Earnings_per_day'] = sample_size[
            ['R01', 'R02', 'R03', 'R04', 'R05', 'R06', 'R07', 'R08']
        ].sum(axis=1)

        # Merge embedded capacity data; default to 65 if fleet_no is missing or unmatched
        df['fleet_no_clean'] = df['fleet_no'].astype(str).str.strip().str.upper()
        lookup_clean = lookup_df.copy()
        lookup_clean['fleet_no_clean'] = lookup_clean['fleet_no'].astype(str).str.strip().str.upper()
        df = pd.merge(df, lookup_clean[['fleet_no_clean', 'total_capacity']], on='fleet_no_clean', how='left')
        df["total_capacity"] = df["total_capacity"].fillna(65)
        df.drop(columns=['fleet_no_clean'], inplace=True)

        # Demand calculations
        df['stages_travelled'] = np.where(
            df['to_stage'] == 0,
            0,
            (df['to_stage'] - df['from_stage']).abs()
        )
        df['demand'] = df['passenger_count'] * df['stages_travelled']
        df['supply'] = df['total_capacity'] * (df['total_stages'] - 1)

        # Create helping DataFrame
        helping_df = df[['service_trip_shift']].drop_duplicates()
        helping_df[['service_no', 'trip_shift', 'service_type']] = helping_df['service_trip_shift'].str.split("_", expand=True)
        helping_df['route_no'] = helping_df['service_no'].str.split("-").str[0]

        ticket_types_free = [
            ('TCH03', 'Disable Person Escort Ticket'),
            ('TCH04', 'Disable Person Ticket'),
            ('TCH06', 'Senior Citizen Ticket'),
            ('TCH07', 'Transgender Ticket'),
            ('TCH08', 'Women Ticket')
        ]

        for col, ticket_type in ticket_types_free:
            helping_df[col] = helping_df['service_trip_shift'].map(
                df[df['ticket_type'] == ticket_type].groupby('service_trip_shift')['passenger_count'].sum()
            ).fillna(0)

        helping_df['zero_value_ticket_pass'] = helping_df[['TCH03', 'TCH04', 'TCH06', 'TCH07', 'TCH08']].sum(axis=1)

        helping_df = helping_df.merge(
            df.groupby('service_trip_shift')['stages_travelled'].mean().rename('average_stages'),
            on='service_trip_shift'
        ).round(2)
        helping_df['zero_value_ticket_demand'] = (helping_df['zero_value_ticket_pass'] * helping_df['average_stages']).round(2)
        helping_df = helping_df.merge(df.groupby('service_trip_shift')['KM'].mean().rename('kms'), on='service_trip_shift').round(2)
        helping_df = helping_df.merge(df.groupby('service_trip_shift')['total_stages'].mean().rename('stages'), on='service_trip_shift')
        helping_df = helping_df.merge(df.groupby('service_trip_shift')['total_capacity'].mean().rename('capacity'), on='service_trip_shift').round(2)
        helping_df['supply'] = (helping_df['capacity'] * helping_df['stages']).round(2)

        # Service-level aggregates
        service_help = helping_df.groupby('service_no').agg(
            service_kms=('kms', 'sum'),
            service_supply=('supply', 'sum'),
            zero_value_ticket_demand_service=('zero_value_ticket_demand', 'sum')
        ).reset_index().round(2)

        # Route-level aggregates
        route_help = helping_df.groupby('route_no').agg(
            route_kms=('kms', 'sum'),
            route_supply=('supply', 'sum'),
            zero_value_ticket_demand_route=('zero_value_ticket_demand', 'sum')
        ).reset_index().round(2)

        # Service DataFrame
        service_df = df[["service_no", "service_type"]].drop_duplicates()
        service_df['route_no'] = service_df['service_no'].str.split("-").str[0]
        service_df = service_df[['service_no', 'service_type', 'route_no']]

        duty_date = parse_date(df["duty_date"].iloc[0])
        if duty_date is not None and not pd.isna(duty_date):
            service_df["date"] = duty_date.strftime('%d-%m-%y')
            service_df["day"] = duty_date.strftime('%A')
        else:
            service_df["date"] = str(df["duty_date"].iloc[0])
            service_df["day"] = "Unknown"

        ticket_types = [
            ('TCS01', 'Adult Ticket'),
            ('TCS02', 'Difference Fare Ticket'),
            ('TCS03', 'Disable Person Escort Ticket'),
            ('TCS04', 'Disable Person Ticket'),
            ('TCS05', 'Luggage Ticket'),
            ('TCS06', 'Senior Citizen Ticket'),
            ('TCS07', 'Transgender Ticket'),
            ('TCS08', 'Women Ticket')
        ]

        for col, ticket_type in ticket_types:
            service_df[col] = service_df['service_no'].map(
                df[df['ticket_type'] == ticket_type].groupby('service_no').size()
            ).fillna(0)

        service_df['TTPS'] = service_df[[f'TCS{i:02d}' for i in range(1, 9)]].sum(axis=1)
        service_df = service_df.merge(df.groupby('service_no').size().rename('TTPS_Check'), on='service_no', how='left').fillna(0)

        passenger_cols = [
            ('PCS01', 'Adult Ticket'),
            ('PCS02', 'Difference Fare Ticket'),
            ('PCS03', 'Disable Person Escort Ticket'),
            ('PCS04', 'Disable Person Ticket'),
            ('PCS06', 'Senior Citizen Ticket'),
            ('PCS07', 'Transgender Ticket'),
            ('PCS08', 'Women Ticket')
        ]

        for col, ticket_type in passenger_cols:
            service_df[col] = service_df['service_no'].map(
                df[df['ticket_type'] == ticket_type].groupby('service_no')['passenger_count'].sum()
            ).fillna(0)

        service_df['PPSPD'] = service_df[
            ['PCS01', 'PCS02', 'PCS03', 'PCS04', 'PCS06', 'PCS07', 'PCS08']
        ].sum(axis=1)

        service_df = service_df.merge(
            df.groupby('service_no')['passenger_count'].sum().rename('PPSPD_Check'),
            on='service_no',
            how='left'
        ).fillna(0)

        service_df['RS01'] = service_df['service_no'].map(
            df[df['ticket_type'] == 'Adult Ticket'].groupby('service_no')['passenger_amount'].sum()
        ).fillna(0)
        service_df['RS02'] = service_df['service_no'].map(
            df[df['ticket_type'] == 'Difference Fare Ticket'].groupby('service_no')['passenger_amount'].sum()
        ).fillna(0)
        service_df['RS03'] = (service_df['PCS03'] * 16).astype(int)
        service_df['RS04'] = (service_df['PCS04'] * 16).astype(int)
        service_df['RS05'] = service_df['service_no'].map(
            df[df['ticket_type'] == 'Luggage Ticket'].groupby('service_no')['luggage_amount'].sum()
        ).fillna(0)
        service_df['RS06'] = (service_df['PCS06'] * 21).astype(int)
        service_df['RS07'] = (service_df['PCS07'] * 16).astype(int)
        service_df['RS08'] = (service_df['PCS08'] * 16).astype(int)

        service_df['EPSPD'] = service_df[[f'RS{i:02d}' for i in range(1, 9)]].sum(axis=1)

        service_metrics = df.groupby("service_no").agg(trips_per_service=('trip_no', 'nunique')).reset_index()
        service_df = service_df.merge(service_metrics, on="service_no", how='left')

        service_demand = df.groupby('service_no')['demand'].sum().reset_index()
        service_df = service_df.merge(service_demand, on='service_no', how='left').fillna(0)
        service_df = pd.merge(service_df, service_help, on='service_no', how='left')
        service_df["service_DSRWOW"] = (service_df["demand"] / service_df["service_supply"]).round(2).fillna(0)
        service_df['total_service_demand'] = (service_df["demand"] + service_df["zero_value_ticket_demand_service"]).round(2)
        service_df['service_DSRWW'] = (service_df['total_service_demand'] / service_df["service_supply"]).round(2)
        service_df = service_df.sort_values('service_DSRWW', ascending=False).reset_index(drop=True)
        sample_size['service_count'] = len(service_df)

        # Route DataFrame
        route_df = df[['route_no']].drop_duplicates()

        duty_date = parse_date(df["duty_date"].iloc[0])
        if duty_date is not None and not pd.isna(duty_date):
            route_df["date"] = duty_date.strftime('%d-%m-%y')
            route_df["day"] = duty_date.strftime('%A')
        else:
            route_df["date"] = str(df["duty_date"].iloc[0])
            route_df["day"] = "Unknown"

        for i, (col, ticket_type) in enumerate(ticket_types, 1):
            route_df[f'TCR{i:02d}'] = route_df['route_no'].map(
                df[df['ticket_type'] == ticket_type].groupby('route_no').size()
            ).fillna(0)

        route_df['TTPR'] = route_df[[f'TCR{i:02d}' for i in range(1, 9)]].sum(axis=1)
        route_df = route_df.merge(df.groupby('route_no').size().rename('TTPR_Check'), on='route_no', how='left').fillna(0)

        route_df['PCR01'] = route_df['route_no'].map(
            df[df['ticket_type'] == 'Adult Ticket'].groupby('route_no')['passenger_count'].sum()
        ).fillna(0)
        route_df['PCR02'] = route_df['route_no'].map(
            df[df['ticket_type'] == 'Difference Fare Ticket'].groupby('route_no')['passenger_count'].sum()
        ).fillna(0)
        route_df['PCR03'] = route_df['route_no'].map(
            df[df['ticket_type'] == 'Disable Person Escort Ticket'].groupby('route_no')['passenger_count'].sum()
        ).fillna(0)
        route_df['PCR04'] = route_df['route_no'].map(
            df[df['ticket_type'] == 'Disable Person Ticket'].groupby('route_no')['passenger_count'].sum()
        ).fillna(0)
        route_df['PCR06'] = route_df['route_no'].map(
            df[df['ticket_type'] == 'Senior Citizen Ticket'].groupby('route_no')['passenger_count'].sum()
        ).fillna(0)
        route_df['PCR07'] = route_df['route_no'].map(
            df[df['ticket_type'] == 'Transgender Ticket'].groupby('route_no')['passenger_count'].sum()
        ).fillna(0)
        route_df['PCR08'] = route_df['route_no'].map(
            df[df['ticket_type'] == 'Women Ticket'].groupby('route_no')['passenger_count'].sum()
        ).fillna(0)

        route_df['PPRPD'] = route_df[['PCR01', 'PCR02', 'PCR03', 'PCR04', 'PCR06', 'PCR07', 'PCR08']].sum(axis=1)
        route_df = route_df.merge(df.groupby('route_no')['passenger_count'].sum().rename('PPRPD_Check'), on='route_no')

        route_df['RR01'] = route_df['route_no'].map(
            df[df['ticket_type'] == 'Adult Ticket'].groupby('route_no')['passenger_amount'].sum()
        ).fillna(0)
        route_df['RR02'] = route_df['route_no'].map(
            df[df['ticket_type'] == 'Difference Fare Ticket'].groupby('route_no')['passenger_amount'].sum()
        ).fillna(0)
        route_df['RR03'] = (route_df['PCR03'] * 16).astype(int)
        route_df['RR04'] = (route_df['PCR04'] * 16).astype(int)
        route_df['RR05'] = route_df['route_no'].map(
            df[df['ticket_type'] == 'Luggage Ticket'].groupby('route_no')['luggage_amount'].sum()
        ).fillna(0)
        route_df['RR06'] = (route_df['PCR06'] * 21).astype(int)
        route_df['RR07'] = (route_df['PCR07'] * 16).astype(int)
        route_df['RR08'] = (route_df['PCR08'] * 16).astype(int)

        route_df['EPRPD'] = route_df[['RR01', 'RR02', 'RR03', 'RR04', 'RR05', 'RR06', 'RR07', 'RR08']].sum(axis=1)

        route_demand = df.groupby('route_no')['demand'].sum().reset_index()
        route_df = route_df.merge(route_demand, on='route_no').fillna(0)
        route_df = pd.merge(route_df, route_help, on='route_no', how='left')
        route_df['route_DSRWOW'] = (route_df['demand'] / route_df['route_supply']).round(2).fillna(0)
        route_df['total_route_demand'] = (route_df["demand"] + route_df["zero_value_ticket_demand_route"]).round(2)
        route_df['route_DSRWW'] = (route_df['total_route_demand'] / route_df['route_supply']).round(2)

        route_df = route_df.sort_values('route_DSRWW', ascending=False).reset_index(drop=True)
        sample_size['route_count'] = len(route_df)

        return service_df, route_df, sample_size

    except Exception as e:
        st.error(f"Error processing {file_name}: {str(e)}")
        st.text(traceback.format_exc())
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()


# Run Analysis Button
if st.button("🚀 Run Analysis", key="run_analysis"):
    items_to_process = available_items if analysis_method == "Process all selected files" else filtered_items

    if not items_to_process:
        st.error("No files available or selected for processing!")
        st.stop()

    progress_bar = st.progress(0)
    status_text = st.empty()

    all_service = pd.DataFrame()
    all_route = pd.DataFrame()
    all_sample = pd.DataFrame()

    for i, (file_name, file_source) in enumerate(items_to_process):
        status_text.text(f"Processing {i+1}/{len(items_to_process)}: {file_name}")
        progress_bar.progress(int((i + 1) / len(items_to_process) * 100))

        service, route, sample = process_file(file_name, file_source, lookup_df)

        if not service.empty:
            all_service = pd.concat([all_service, service], ignore_index=True)
        if not route.empty:
            all_route = pd.concat([all_route, route], ignore_index=True)
        if not sample.empty:
            all_sample = pd.concat([all_sample, sample], ignore_index=True)

    st.session_state.all_service = all_service
    st.session_state.all_route = all_route
    st.session_state.all_sample = all_sample
    st.session_state.analysis_complete = True

    progress_bar.empty()
    status_text.empty()

    st.success(f"✅ Analysis completed successfully! Processed {len(items_to_process)} file(s).")

# Display results if analysis is complete
if st.session_state.analysis_complete:
    st.header("📊 Analysis Results")

    if not st.session_state.all_service.empty:
        st.subheader("Service Data")
        st.dataframe(st.session_state.all_service)

    if not st.session_state.all_route.empty:
        st.subheader("Route Data")
        st.dataframe(st.session_state.all_route)

    if not st.session_state.all_sample.empty:
        st.subheader("Summary Data")
        st.dataframe(st.session_state.all_sample)

    timestamp = datetime.now().strftime("%d-%m-%Y_%H%M%S")
    output_excel_name = f"ETM_Analysis_PS_{timestamp}.xlsx"

    try:
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            if not st.session_state.all_service.empty:
                st.session_state.all_service.to_excel(writer, sheet_name='Service_Data', index=False)
            if not st.session_state.all_route.empty:
                st.session_state.all_route.to_excel(writer, sheet_name='Route_Data', index=False)
            if not st.session_state.all_sample.empty:
                st.session_state.all_sample.to_excel(writer, sheet_name='Summary_Data', index=False)

        st.download_button(
            label="📥 Download Full Analysis Results (.xlsx)",
            data=excel_buffer.getvalue(),
            file_name=output_excel_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Error creating Excel report: {e}")

# Sidebar Instructions & Notes
st.sidebar.header("ℹ️ Instructions")
st.sidebar.info("""
1. **Select Data Source**: Test immediately with built-in sample data or upload your own `.csv` / `.bz2` files
2. **Select Analysis Method**:
   - All Files: Process every selected file
   - Date Range: Filter by start/end dates
   - Monthly: Filter by month
   - Weekday/Weekend: Filter Mon-Fri or Sat-Sun
3. **Run Analysis**: Generate Service, Route, and Summary tables
4. **Download**: Export the full multi-sheet Excel report
""")

st.sidebar.header("📝 Notes")
st.sidebar.warning("""
- MTC Fleet Lookup (3,390 vehicles) is built into the app; missing/unmatched fleet numbers default to a capacity of 65
- Works with any filename (`sample_etm_data.csv`, `.csv.bz2`, etc.)
""")
