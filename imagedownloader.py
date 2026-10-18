"""
Images to ZIP + PPT maker - Streamlit app.

Three ways to use it:
  1. Upload Excel  -> downloads every image column, files named  <StoreCode>_<ColumnName>[_n].jpg
  2. Paste links   -> one URL per line, get a ZIP
  3. Make PPT      -> upload Excel, get a PowerPoint: one slide per outlet
                      (Store front | Before deployment | After deployment)
                      with a blue header: client store code + store name

Run locally:
    pip install streamlit openpyxl python-pptx pillow
    streamlit run app.py
"""
import base64
import io
import math
import mimetypes
import re
import zipfile
from datetime import date, datetime
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse
import urllib.request

import streamlit as st
from openpyxl import load_workbook
from PIL import Image, ImageOps
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/124.0 Safari/537.36"}
URL_RE = re.compile(r"https?://[^\s,\"'<>]+", re.I)

# Channelplay logo (small fallback, cropped from your slide). Upload a sharper one in the app if you have it.
DEFAULT_LOGO_B64 = "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAMCAgICAgMCAgIDAwMDBAYEBAQEBAgGBgUGCQgKCgkICQkKDA8MCgsOCwkJDRENDg8QEBEQCgwSExIQEw8QEBD/2wBDAQMDAwQDBAgEBAgQCwkLEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBD/wAARCABcAYoDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD9TsfwsM+/rWdr2uWPh7TLrVtRnSK3t0LuzHAAFaeQTXin7XZC/ArxHm4eEgR/OhwfvDiurA4eOIxEKUurS+9nn5niKmEwlStHdJs801T9uvTLXUJ4dM8GPd2sbsiz+Zjdg+lVD+3jEf8AmQZP+/tfI0ZDImI/LDIuPfjrRmv2eHAmXRgpSV36s/nqr4h5xzyUZKyPrr/hvKE/8yBJ/wB/qUft5RZ2nwBJn/rrXyKtLucMVdQR605cD5ZytqP4sj/iIudN+7JfcfXDft5w9/h/IP8AttQP28IF3Z+H8h29f31fJGCUIGBTH3MAmPuZLH1o/wBSMq5E2n97Kh4i5xKWrVz66P7eEDNtHw/kBHrNWhoH7cukajq8NlqvhN7C2kcIZzJkAmvLPgr+zPefFDQ28Q3+pPZwfdg+T71cB8W/hlffCzxSPD99MZoiwa3lIxurxo5Jw/WxE8HBvnXmz3pcR8S0aMcbUX7uTXQ/TbSdTtdXsINQspA0U6B1Yc5Bq5XmH7O1zdXPwr0OS5kMknlkEn0r1Dmvy7GUlhq06cdk2ftWXYmWMwsKst2kL060tM9qfXMlod3UKQnHJpaYSqk8/eosA4soIGevSlqMjfkH7p6H0qSiwBRRRRYAoooosAUUUUWAKKKKLAFFFFFgCk3DOO9LTeRwetFgHUmRnHc00KFOaRZQwBXJ3ZxRYB5OBknFA65zTAG+XHAGcinZ4xRsK1xc00t79KUc8VEeDsjfaFPIxTswk1Em3DGc8UtMyDginDpSsx6W0FoooosAUUUUwCiiigAooooAKKKKACkwKWikwGbcdBXif7XgVPgX4jJGR+74/wCBCvb68P8A2vsn4F+IvpH/AOhCvQylf7bS/wAS/M8biD/kXVn/AHX+R+ecZfy4hK28BFx27Uqq5OI4ZJCegjXcf0pq7vLT08tf5V9SfsdeH/B19aatfazBavdKNoE5B2g9cA1++5jmH9mYN4lpu1tD+Xcqyp5vjPYKSje+58tKSScAgjqCMV6j8B/hDH8XNdurW4vDBb2g/eEdqo/HvTPDOl/EzUY/C13G1mhy8aD5Ubvg10P7PPjDxb8O21jxJovhKTV7GQAzuG2hcehrmzHHVcRlTr4d8s5LRM7spy+lQzVU8R70YvW2pj/HL4Qt8JtfhsYrsz2txzGx715o7bFO2RYw/TJ616v8QvHeuftD+NLCzsdNWzmjYpFA78Z7817l8Lv2MNAhSHVfHkov3OH+yK3yo3bkda8+WfxynL4f2i/3jWy3/r8D0v8AV2efZlP+yo+4nv0H/s1fHjwppHgaHw34puDpslgD5bSJt84eorxP9p74l6V8TfGVtJorgWemviKfH+tPeuu/bF0HStB1XSNO0iyS1iiXaBGNvA/nXziArHe7ZIYdqwyLKcJXbzmne8k9H0OzP86x2EjHJKtrQa1R+lX7OZ3fCnRCBg+XXqBz2715d+zic/CnRGH/ADzr0/hSMnn+EelfkOaf75Vfm/zP3jItcupeiHqwI60hkA7N+VeB/tJfta+Df2e9PaCe2XWdemGYNOikG5z7+lfIM/7Tv7c/xVb/AISH4b6QdE0uUlo4pbPccemSK89anr26H6dCQEZw34ivNvjv8ZdM+CXg8+K9RtzMHlSIL6ZOK+Hvh5+13+1p4A8Y6bpPxo0NtWttVmES7bbyljBOM5xXuv8AwUBvP7R/Z9tNQ8vaZ5reXHpuIOKdh2PpfwP4rtvG3hfT/EtkmyG+iWQL9RmuhrzT9nYY+DnhnGObSP8A9BFel0CCiiimLUKKKKBhRRRSsAUUUUWAKKKKLAFMH3mPan0w5OfSiz6AZfiTWk8O6Be63MuUsommf6KM15f+zz+0Zo/x+g1mfR7M2y6TMYWz3IOM13HxWH/FufEOf+fGX/0E18g/8Ew8f2Z444A/4mD9B/tGl2HbQ+6QCRtY7vU0o4yDSDOeCMU4YzQrr4haCV5l8U/jRpXww1jRdI1G2MsmsSbIyPrXppJXkdK+Tf2xsHx14BOAf9IPX617GTYSnjsWqNXZ3/I8DiPMKuXYGWIor3k1+LPqmwuxe2cV7GvyyoHUemRVvcKyvDJ/4kdln/ngn8q09wJry69oVHDzPXw03Uoxm92kLvXOM0B1JxTSuTnNLgKc1m072R0PpYfSZFJyaDwOlUkJuyuG8HpS7qaABzjmkLkUm1eyBtJXH5FLTMkjNPoHuFFFFMAooopMArxD9r7P/CivER/65/8AoQr2414h+12yj4G+IVJyCI8/99CvQylXxtL/ABL8zxuIWllta/8AK/yPzy3BYYsnGY1x+Ve6/Az4HeKPGegan4ij1K60qIRM1tKuR5uB0xXiOlXFnZXNtNqUBuYo1WTaO+B0r9C/2ffG9r4k+Gh1m10xLK1ghbZbAdAo5/Ov2LizMq2CwcVRje9tT8B4KyvD4/GT+sSta+x8X+Avhw3i7x3NofiTUhZxWtwUuWlb5jz79c194W/wz8K+HPh1d+HPD2nokElo2GHVzt65r5Ck/sj4pfHexv8AwjZvC09yRfWyvtA2Hr79K+i/2k/ibd/C/wAEWdnozhbu7UIgzyFXqK+YzyriMfiMNQovlb15e1tT7HhulhMswmJxNaCkldX9dND4k0rRdb/4WDH4d0dZINTN6yxhSQQu7k5+lfpZ4T05fB3hGztdV1Ay/Zolaa4lbqcc5Jr5Z/Za8LRfELx5efFDU7DyUswBCCOrHqfzrsv2wPirJ4f8NjwdpmPO1H5ZDG/zIP6VnxBKpnOOp5dSWsdzThdU+GstrZrVl7s/hXzPNP2ytV0/XdX0rUNMuUuYGyC6HI4r5uXH93jcKvT67fT6UmlzXDyonKbySR61QXt/vCv0fLMv/s3A/V+yPynOcyeb5i8W/tNH6Ufs5HPwp0THA8s1q/G/4mWXwi+Gus+OLtQ7WcDeUhONzkcflWV+zhx8KNE/65188/8ABTXxj9i8A6P4GAZf7auFDMD1Gelfz/mn+9VPV/mf1DkH/IupeiPJP2SvgtqP7TnxD1L40/FeaW/02K4Z4YJSdsgzlQPTFfpHpmiaVolhFp2k6ZbW9tAoSOJIwABXl37KHhTTfCXwP8OWmnQKnnW6vKQOWOK9jwPSuGOx7Mnqc74l8C+FvF0UcPiLRba6WFg8R2AMhHoa+b/+Ch0UNn8CYLSNdsSXUKKB2AIxX1gcgV8pf8FGJ0tvgaZ5mCoLqJn74AIzQxI9k/Z2Oz4OeGxg5FpH24+6K9K3Cvy6m/bD+LOoeENE8D/ArRryS3021QT3a25IOF55xXc/sp/tzeMNQ8f/APCu/jY00Qvm8u2lmhKGJxxg/U0kx2P0M8xdxXPIoDg+vPtXIfEz4l+Hfhb4MvfGfiC4RbW1j3opPMhxwB9a/N/xR+3X+0P8TfE1xc/DPQ7uz0uByILaK3LjA6EsBzmncVj9T9454PHbFKGyMgH8q/PLRP8AgoR4ii+FGr6N4yt59E+Itkypp/2i3IW9HfgjtX1L8Lvil4m8Tfs//wDCw9VKPqqWcszbVwGZVyOPrRcLHs+4eh/KjcAQOefavzd0n/goh43uvAGt6TCkmoeN7i7eDTY7e23CABscgDmuL0j9tv8AaT+G3iy2m+I1ndfYZJI/tMEtuVUhj1DEUXCx+q+RkjPSjd7H8q5Hw543tfGHw8g8baMwWO7sWuYsc7W25/nXwXbf8FCfiVpl14w8Lahpx1DV1uGg0QwQbiuDjkAc0XCx+kO7v/KlyK/J9P20v2ovAGvWuqeOLG6t7eaVXliltyEljJ9TwOK/Sn4R/FDRfiz4E0/x1o7eXBdxbpY2P+rYdaLhZo7beMkc8deKaXAHXOfSvgj9qL9uvxRpfjKX4Z/Baylury2lEM1zBH5gkdjjAx05rhbD9rP9qT4C+JdMf44+FpjpmplWEYUsuxv4i3bijR6thZdT9Bfiuw/4Vx4i5/5cZf8A0E18g/8ABMU7NK8cMRx9vf8A9CNfTWqePNB+JXwN1Dxl4emElpe6dI6kH7rbORXzH/wTJLppnjx3bOL5zt/4EamXvWURJ8qblsfdgA6nNAYDjP6V8xzftG6xoHxK17SNdukGk6Wh8qMAbnOOBXmevftIfHLXLu48QaLo11a6RE7eTH5Byo7H3zX01LhrF1fiaSte7Z8fX4wwFBe5GUmux90ZBGOefUV8mftiZ/4TnwF14uT/ADrsv2ef2jj8RoR4b8Txm212AfMzDHm/h2rjf2wiJPG/gRc4BuOD75rpyLBVsDm0aNZapP8AJnBxJmdDM8ldfDy6x/M+oPDgP9hWJH/PBP5VpFlBA6Z6VzUeu2HhrwbBqmqTrHb21qrs5OOi18i+MP2oviX4w1+ex+HtjOLBXKwlIyd2D3NedhMmrZnWm4WSTer2PXxmf4XJ8ND2icpNLRb7H3DnsOtNBIBLg4+lfGNt+1n440Hw7e6V4u0iez1qJP3Fw0ZCk/1rsNA+M/xN1/4L3XjfT2WXVLfJKheMD2rWrw1i6C55tWva99zChxpgatoKLv2tqj6fBHVuBSllC7u1eR/s9fFz/haPhM3Go3C/2vZHZfxAfcbtXZ/ETxbD4N8Hah4hlnVDaxMyM3HPavJqYKtRxP1aS969j3aGZ4evhPrkX7p06MSTuJ/KnZz1HSvB/wBmn4ieOviVaahr3iaUC0EhW3j2YwM8GveARjpSxmEeArOlPddi8vx8Myo+3gtGL1FOpvQcDrTq5Ud4UUUUwCiiikwEPSvEf2vFA+BHiNsdBH/6EK9vrw/9r9iPgZ4iX18v/wBCFehlSbxlJL+ZfmePn8oxy6q5bcr/ACPzyQlYEMZEZEQ5PPavrv4Q67d6F+zTqWoWk4jnSN1U4+6D1r5CP+oX/rkP/Qa+lLC+i0b9lQeYW23O7aQOn1r9o4lgp0aNC27j+B/OvDNZ0MRVrLbX8TG/ZF0e51f4lTa2Ity2Lu8jjjBb1qH9qnxNeeMfiinhy3l81LB1jiVWyGLcGuu/ZcW38JfDXxT43vJvJa6jIVz9OMV598DPCE/xN+LM+oXMjG0sZ3nnkY5zg5HNeQpU6eZVsfV2pJJerR9CnUqZbRwFJ61G2/kz6n8AQ6f8D/gydW1Xy0kjt/MkUkAsxHA96+GfHvjDUPHPie98SX8rl55D5aFshV7V7H+1h8UY/Emsw+DdKlYWWk/JL5bfLKelfPfOPp0r0OE8s5ZTzHFr356+nY8vjHOIz9nluHf7umvvYu5s7lxvHRqRev1YUo74po6g/wC0K+yq3dOV+x8JR1qxS7o/Sn9nD/klGif9c6+dP+CnXhZrnwLovjISOTo84IRUz37+lfRf7OH/ACSnRP8Arma1fjV8PIvif8M9d8FTMm+/t22Oy52sBxiv5qzT/eqnq/zP66yBXy+l6I5r9k7xXp/i34HeHL2ymQvHbqs6A5KNivZMjpmvzY/Y9+NkP7PPjzU/gf8AEySSwha4ZIrmbIU84XGe1fozY6lZalbx3mnXcN3BKAySQuHGD06VwR2PXtrqXWI6Zr5O/wCCj8Yk+A7w7NzS3MaZ9iRX0h4o8b+FfCEKP4i1m3tRM4SONpBvYn/Z6184f8FDLmGf4FQXEXzRTXMLKfYkYqrAj0H9kzwH4e8J/BTw7FYaZaLNcW6vLKYRucketfJX7cOlWXhv9ozwhq2lW1vbzXV1Dv8ALjC/xD09a+2v2elI+D3hgMv/AC6x4P4Cvi3/AIKAAf8AC9vBmIiT9rh5z/tCpasPqdB/wUK8T3esQeA/htbM8La6sAeQNwobGcjvX1X8B/g14Y+Efw60rwzo2nWpmW3SW4uXiDNM7DJOa+T/APgoBoE+lw/Dn4mPue10dbczYHQADNfZXwh8eaD8SPh9o3ifw/dx3FrPaxgFGyysFwQw7UgPm7/gop8FdB8S/CxfiFpun28Gs6BKpRooghlVjzkj0rsfgKv/ABiEUB2smmTFye52c1gf8FE/i1o/g74QHweZopNV8QzKltEHG9Ap5JHUVv8AwK+b9kR/M4zpc2T7+XzQHQ8N/wCCaHw48P3Vx4o8dXNpBNdpfTJGssQfbljyM9K9v/bx8A6R4m+BepancQW0M2lsJYysQDMT2zXn3/BMwFPCHiWPaNv2+bkf75r2n9s/B/Z78SDbuO1cUw6lD9lQt/wytpO4nIsJwcnPRTXzL/wT/wDCWg658d/iLq+saVFdXOkXT/Y3lUMItzc8Gvpr9lQMP2WNLAG1vsNxn8q8C/4Jz5HxW+LKhMMLw98/xUh3Paf2+NB0a7/Zm8WahcaXatc2CI9tKIwGjO4Dg/jXmH7M/i6Xwl+xbrviWS6MX2SzaNWUfcLJjj869e/b02/8Ms+N9wziKP8A9DFeO/sw6GfE/wCxbruhQWxupJrNnaHHLbUzx+VMSZ89fsdfHL4XfCPVNb8RfEW0bUNRvbmSS2keDzSuTnOTXufx7/bK+AXxd+G2q+FbjTZZriWIi3Z7bLI2OMHtXK/8E+/Dvwg8d2et+EvGfhK2m1e2u5EhN0AHIUnI2nmvrDxp8F/2cfAPhu+8V674G063tbGNmcuoUtjsM9al66MrmS3Pnv8AYg8RX2rfsz+OdGuXkMGlwTC2DjG1SpxVn/gmQVbRfHY2net64LevzGvYvAOu/BrxP8CvE2rfBvTI7OwuLWYXcKDlWCnqa8e/4JmHGjeO1D5H21znH+0a0WkkomVT+HK/Zlu08Caf45/al1CHUFIt7dg7x54kI5GRX2ZDoOlx2gtF0+2EQGwqIhjFfHejeMNP8IftT6mdQkWO3v2CGRjgIe1faCXEL2y3KSAoRuDA8EH3r6niKdZSpK7tyqx+f8JQwdWNZpLmTd7nx78T/D6fCz9oDSNb0CIQwau4UQKMDOeTWh+1w4m8VfDl3O3zJQx+uRVP4v643xB/aA0Xw/pEiyjQ5FZ2Q5HPXmrv7W8Yi8W/DqJ13FJAP1r6PBRl9YwbqfG4yv3tbS58zmKp/Vsa6K/d80bdt9bG1+1Z4gvtP+GOi6JbO0a6kkaOyntgV6B+zv4G0TRPh7p9xb6dALi5jDyuyAljXnn7VWg39z8NdG1y3B8vTUjZhjPBAr0/9nfxPpniD4b6aLSdZJYIwsyg8qfpXiY1uOTx9j/M+a3rpc9/LVTq54/rK+yuVP0Wxzn7U3g3Rb/4cXuotp8P2mzUlXVACM1z/wCyVpcOq/B+40y4UMkodDnnmup/an8S6fpHw11CxuLxIp71CsKk/M34VifsYLu+GwKjG5ycUU6ldZHKU27KSt95NWlhp8RqnTirtO9vQ8n+GOsXPwV+PWo+GLhTFaarL+9U9GJPy12H7XvjK81gaR8PtGkMg1CRftUanBIJ+Wtb9rn4essFj8StAtf9O0mQPc7BzKOMflXnXwB0jV/jf8VB458SK32fQ0VVQ9GOOK9ahPD4iEM6ktYJqS89l/XoePXWKwuIlw/HaTTT8t2fU/wg8FjwT4I0zSWjVJ1hUy4HUkV3VNjQRoq46cAe1P74FfntetKvWlVnuz9ZwmHhhKMaMFokGeKdTQMinVj1OkKKKKYBRRRSYBXh/wC2AVHwK8RMe3l/+hCvcK8P/a+cp8DfEbAZx5f8xXoZSr42lb+ZfmeLxD/yLa3+F/kfnjg+TEx6MiD9K90+IWo3WkfAnw3oMcgCajkMAM5zXhYD7bcMcs/l49hxXr/xG1K1u9A8G6HBMzPbSIz4XJAyM8V+8ZnSVWth5NXSv+R/NOXVFToVujuvzOg8X6jL8PvgDofhOF/PufECkyY4MXpWx4PvtN+BXwRk8Qzw7td8TRlI0Jw2MYz7Vxfim8/4Tr4o2OkxYm0bRYI5WlJwGKjJGO1ch8V/HV1408RsVby9PsgIbeAfdXbxXi0steJUaMtm+aXnrov67HtVsxjhVKtTeqVo/qcjdXM97cteTzM8ru0jFjkndzioyfSkGMMx++Og9aVRkZNfbQhGC0+E+EqzcnzSdwJAxQflIz3YU3JNODAGP0VuSamrK1OTfYqhFqrH1R+k/wCzh/ySjRB6x16iEJGc4ry/9nJWHwo0Tf8A88zivUjxxmv5szPXF1PV/mf1xkbccvpeiPm79qT9jvwt8d9Nl1jSdml+KYlzFeRjBcjsa+RtL+Gf7ePwBsm0zwRrU8dnK+0iUmfdjoRnpX6k9iKaY1YCNkDr/tc1wJWPXvdH5neEf2XP2qPir8Q9P8e/FfWrlgZEkuD5pVMKcjCdBX1d+1p8H/FHxJ+Dtn4L8K2f2u7t2gBQnGQmMnNfQm3Bzuwo7CnYzxzimBxPwb8Oal4U+HGiaBqybLmygSN1z90gDIr5p/a3/Z3+I3xT+K3hrxJ4UsBPY6dPG87FsbQDzX2Z6gjikYDIJB/A0DR5h8Vfg1pfxX+Fg8Eawi+atmkcbH+GQLj+dfAOk/A79tX9nTUtQ0n4T6tJb2epS/MW/eoy5+XaOi/hX6nYHXPFJtVuMA4/vDNAXPzguf2Jfi/8QfBuu+P/AIr6hNrXje8MbabE0hCw+uF7V9ZfCv4b+IvDf7Oh8BarbeVqZsZYTHnPzMuAPzr2323c0uCep49KAbufMH7EPwW8a/B3w7rWneMtP+xzXl3JLEm7OVLEg16Z+0p4J1z4hfCDW/Cnh2HzL68CiFM/er1PkHGOPXNGOgA4oEeRfATwDr/gn4DWHgjW4dmpwWk0UiZ6MwOK8j/Y2+Anjz4R/EPx/rvi3T/IttcuDJaPuz5g3Zr65x1BHHrQB/E3Gfegdzx79q7wD4h+JvwI8TeCvC9r5+pakirBGTjcQwNYX7HPws8U/Cv4Vw+GPGFosN4gAcdQwx0r30qcYB59aUrng9KAufA/7Qv7IfxT8O/FC5+Mf7P+o/Zb68zJLBF8qRkc52j1rzK5+Ev7cX7RiWmgfE7XnOlW0uCUXyQFz827+9x61+omB0xxTVRVXCKqD/ZGKVkwvqeNeCPgFo/ws+Cd38OfCNqEu7uxeO4kzzLKVPP5151+xF8D/HPwa0nxZa+M7EW51C6eWDnO5SSc19VDPTf+lR3RHkONx5U04q00Z1nanL0Ph+b4eaR8Rfj74l0HUC8Tuu+GRTyrjpV7UfD37Wvh6EeENP1kz6aC0UL7csI+gGfpXQ/DgOP2o9fAAVFHBNfWQCsoOwMc9T2r7fMM5eDqQjyRkuVbrZ+R+aZNkFLNKUqsakoSbd7O1/U+dv2b/wBna88Ezz+LfGJaXV7s79jtuKH61d/aM+FXirx/4o8J6j4esxLDpcxe5YnGwZr6AAwM4zikKgjnn2zXgLOsT9c+tt+8vuPqo8N4OOBWAs+Xdvq7anM6h4QtPEfgz/hGdbhE0UtuIpAexx1r5KvvhF8e/g1q12PhjqDPZ3MhZXI3DBPTHtX26pA4AxTWRXwpQFR3NGAzirgZSTipRlunsGZ8O4fMeScZOMoaJrc+KtT+APxq+J2k3Gu/ELV2udSVP9EgHyqOPSvc/wBmn4e+Ifh34JOieIYRDdBicg54r2LCr9PQVBcSLFC0u7CoCxJ9BWuMz+vjqX1blSg+iWxz4Hhihllf69zuUl1Z4h+1h44j8MeBhotqytf6yfJt4+pPrWn+zL4FXwZ4Btnli23V6PNmOMEk814T4pnvfjL+0ZFp+mbr6w02QYX+GPb1r7SsLWOzs7e1hjCJEgXAHTArrzGTy3L6eDfxS1kvyucOU01nGa1cwt7sdIv8ywN2BmlOQPrTSSwKnjFOLYxnmvlrJan3Xw7i5OadTSaUc0JWQW6i0UUUwCiiikwG7h2rkPil4DtviL4K1LwleSbY75Pve46V1iqAGbuakHStaVSVGopw3Wpz1qMcVRdOez0Pzi1H9kz44WN7Ja6b4SiuraGT9zMbgAlQeK2z8C/2jTqlnqp8C2xk0+MxwAzqQMjHIr9AdoPc/nRsHqfzr7D/AF5zFxUXGLt5f8E+B/4hnlN5WnPXzX+R+eln+zx+0TZQXzx+D4jdX7EvILkDb7Vj/wDDLXx/UOzeBoneQ8n7SOK/SPYPU/nRsHqfzrSPH2ZRvaMNfJ/5mc/DDKZpJznp5r/I/Nw/sr/HzcAPA8Pyjr9pHNIP2WPj+uf+KHh59Lkf41+kmwep/OjYPU/nVLxAzRR5eWNvR/5kf8Qsye1uaf3r/I/NsfssfH0dfA8R/wC3kVb0j9k7423mqW0GreEIrazMimaQXAJAzX6M7B6n86NoHc/nUz4+zOceVqP3P/Mun4X5PTmpqU9PNf5HOeA/C8fhDwxZaBDylrGFz7966PBIp1FfF1akq03Uluz9Dw9COGpRow2Wg09BRtbpup1FZm43BA2rx70pyfalooATHGDzSAMq4JyadRQAnOPem7TtG75j+VPooAT8KTH8OOKdRQAmBjGKCD0HFLRQAgB6HmkOScFcinUUAJjJz0NIBg7QMD1p1FADTnptzSbSMAcCn0UmrgNwQORmmyKWQr1LDFSUU1oJpNWZ4f4X+B+r6H8ZNT+Ic+p+bbXg+WLGK9rVGA27cVLRXRicTPFtSqbo5MFgaWATVLqMIOOKFXH1p9Fcz1VjsWisMwT2xS4zwRxTqKYrERjIyq8CsbxZpWoaroF5p+nTGG4uI2RH/uZFb1FVSk6TUokVaaq03Tb0Z4h+z/8AAa7+Fd5rOsaxqn2++1aTfkrzEM9Aa9rO4AYqSit8Vi6mMqurVd2zlwGX0suoLD0dkMKnsPrSlc4z2p1Fctjueu43HHNKOKWimAUUUUAFFFFJgf/Z"


# --------------------------------------------------------------------------
# Excel / download helpers
# --------------------------------------------------------------------------
def parse_links(text: str) -> list[str]:
    seen, links = set(), []
    for token in re.split(r"[\s,]+", text):
        if re.match(r"^https?://\S+$", token, re.I) and token not in seen:
            seen.add(token)
            links.append(token)
    return links


def clean(text: str) -> str:
    """Make a string safe for use in a file name."""
    return re.sub(r"[^\w\-]+", "_", str(text).strip()).strip("_")


def expand_gallery(url: str) -> list[str]:
    """Gallery links (.../gallery?i=<base64>) hold the real image URLs, one per line."""
    try:
        i = parse_qs(urlparse(url).query)["i"][0]
        i += "=" * (-len(i) % 4)
        decoded = base64.urlsafe_b64decode(i).decode()
        return URL_RE.findall(decoded) or [url]
    except Exception:
        return [url]


def cell_urls(cell) -> list[str]:
    """All direct image URLs in one Excel cell (hyperlink target, plain text, or gallery link)."""
    candidates = []
    if cell.hyperlink is not None and cell.hyperlink.target:
        candidates.append(cell.hyperlink.target)
    val = cell.value
    if isinstance(val, str):
        candidates += URL_RE.findall(val)  # plain URL text or =HYPERLINK("url","text")

    urls, seen = [], set()
    for c in candidates:
        for u in (expand_gallery(c) if "/gallery" in c and "i=" in c else [c]):
            if u not in seen:
                seen.add(u)
                urls.append(u)
    return urls


def read_workbook(file_bytes: bytes, sheet_name=None):
    wb = load_workbook(io.BytesIO(file_bytes))  # not read_only: hyperlinks need full mode
    ws = wb[sheet_name] if sheet_name else wb.active
    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
    return wb, ws, headers


def build_jobs(ws, headers, code_col, image_cols, rows=None):
    """One job per image: (url, file stem, store code, column name)."""
    code_idx = headers.index(code_col)
    img_idx = {name: headers.index(name) for name in image_cols}

    jobs, used = [], {}  # used: file stem -> url
    for row in range(2, ws.max_row + 1):
        if rows is not None and row not in rows:
            continue
        code = ws.cell(row=row, column=code_idx + 1).value
        code = clean(code) if code not in (None, "") else f"row{row}"
        for col_name, idx in img_idx.items():
            urls = cell_urls(ws.cell(row=row, column=idx + 1))
            for n, url in enumerate(urls, start=1):
                stem = f"{code}_{clean(col_name)}"
                if len(urls) > 1:
                    stem += f"_{n}"
                if stem in used:
                    if used[stem] == url:
                        continue                      # exact same photo repeated in another row
                    stem += f"_r{row}"                # same name, different photo -> keep both
                used[stem] = url
                jobs.append({"url": url, "stem": stem, "code": code, "column": col_name})
    return jobs


def ext_for(url: str, content_type: str) -> str:
    suffix = Path(urlparse(url).path).suffix
    if re.fullmatch(r"\.[A-Za-z0-9]{2,5}", suffix or ""):
        return suffix.lower()
    return mimetypes.guess_extension((content_type or "").split(";")[0]) or ".jpg"


def fetch(job):
    try:
        req = urllib.request.Request(job["url"], headers=HEADERS)
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = resp.read()
            ctype = resp.headers.get("Content-Type", "")
        if ctype.startswith(("text/", "application/json")):
            raise ValueError(f"got a web page instead of an image ({ctype})")
        return {**job, "name": job["stem"] + ext_for(job["url"], ctype), "data": data, "error": None}
    except Exception as err:
        return {**job, "name": None, "data": None, "error": str(err)}


def run_jobs(jobs, zip_name):
    progress = st.progress(0.0, text=f"Fetching 0 of {len(jobs)}...")
    results = []
    with ThreadPoolExecutor(max_workers=4) as pool:
        for done, res in enumerate(pool.map(fetch, jobs), start=1):
            results.append(res)
            progress.progress(done / len(jobs), text=f"Fetching {done} of {len(jobs)}...")
    progress.empty()

    good = [r for r in results if not r["error"]]
    report = [
        {
            "Store Code": r.get("code", ""),
            "Column": r.get("column", ""),
            "Link": r["url"],
            "Status": "Added" if not r["error"] else "Failed",
            "Details": r["name"] or r["error"],
        }
        for r in results
    ]
    out = {"bytes": None, "report": report, "good": len(good), "total": len(results),
           "name": re.sub(r"[^\w\- ]+", "_", zip_name.strip() or "images") + ".zip"}
    if good:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_STORED) as zf:
            for r in good:
                zf.writestr(r["name"], r["data"])
        out["bytes"] = buf.getvalue()
    st.session_state["zip_result"] = out


# --------------------------------------------------------------------------
# Date filter
# --------------------------------------------------------------------------
def parse_date(v):
    """Excel date cell -> datetime.date, or None. Handles real dates, 2026-10-17 and 17/10/2026 (day first)."""
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    t = str(v or "").strip()
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})", t)
    if m:
        y, mo, d = map(int, m.groups())
    else:
        m = re.match(r"^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})", t)
        if not m:
            return None
        d, mo, y = map(int, m.groups())
    try:
        return date(y, mo, d)
    except ValueError:
        return None


def date_filter_ui(ws, headers, key):
    """Shows the date filter. Returns the set of Excel row numbers to process, or None for 'all rows'."""
    named = [h for h in headers if h]
    NO = "(no date filter)"
    default = pick(named, "check-in date", default=-1)
    col = st.selectbox("Date filter column", [NO] + named, index=default + 1 if default >= 0 else 0,
                       key=f"{key}_datecol")
    if col == NO:
        return None
    idx = headers.index(col) + 1
    dates = {r: parse_date(ws.cell(row=r, column=idx).value) for r in range(2, ws.max_row + 1)}
    valid = [d for d in dates.values() if d]
    if not valid:
        st.warning(f"No readable dates found in “{col}”, so the date filter is off.")
        return None
    lo, hi = min(valid), max(valid)
    c1, c2 = st.columns(2)
    start = c1.date_input("From date", value=lo, min_value=lo, max_value=hi, key=f"{key}_from_{lo}_{hi}")
    end = c2.date_input("To date", value=hi, min_value=lo, max_value=hi, key=f"{key}_to_{lo}_{hi}")
    allowed = {r for r, d in dates.items() if d and start <= d <= end}
    unreadable = sum(1 for d in dates.values() if d is None)
    st.caption(f"{len(allowed)} of {len(dates)} rows are inside {start:%d %b} – {end:%d %b %Y}"
               + (f"; {unreadable} rows have no readable date and are skipped" if unreadable else ""))
    return allowed


# --------------------------------------------------------------------------
# PPT builder
# --------------------------------------------------------------------------
BLUE = RGBColor(0x00, 0x70, 0xC0)
GREY = RGBColor(0x7F, 0x7F, 0x7F)
LIGHT = RGBColor(0xF2, 0xF2, 0xF2)
SLIDE_W, SLIDE_H = 13.333, 7.5       # inches (16:9)
BAR_W, BAR_H = 10.7, 0.65            # blue header bar (logo sits to its right, on white)
MARGIN = 0.35
COL_GAP = 0.25
COLUMNS = [("front", "Store Front"), ("before", "Before Deployment"), ("after", "After Deployment")]


def prep_image(data: bytes, max_side=1400):
    """Fix phone-photo rotation, shrink and re-encode so the PPT stays small. -> (jpeg bytes, (w, h))"""
    try:
        im = ImageOps.exif_transpose(Image.open(io.BytesIO(data))).convert("RGB")
        im.thumbnail((max_side, max_side))
        buf = io.BytesIO()
        im.save(buf, "JPEG", quality=82, optimize=True)
        return buf.getvalue(), im.size
    except Exception:
        return None


def boxes_for(n, x, y, w, h, gap=0.12):
    """Where each of n images goes inside a column box."""
    if n <= 1:
        return [(x, y, w, h)]
    if n == 2:
        bh = (h - gap) / 2
        return [(x, y, w, bh), (x, y + bh + gap, w, bh)]
    if n == 3:
        bh, bw = (h - gap) / 2, (w - gap) / 2
        return [(x, y, w, bh), (x, y + bh + gap, bw, bh), (x + bw + gap, y + bh + gap, bw, bh)]
    rows = math.ceil(n / 2)
    bh, bw = (h - gap * (rows - 1)) / rows, (w - gap) / 2
    return [(x + (i % 2) * (bw + gap), y + (i // 2) * (bh + gap), bw, bh) for i in range(n)]


def _rect(slide, x, y, w, h, fill):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.line.fill.background()
    shp.shadow.inherit = False
    for ref in shp._element.xpath(".//a:effectRef"):   # also drop the theme's shadow reference
        ref.set("idx", "0")
    return shp


def _text(shape_or_tf, text, size, color, bold=False, align=PP_ALIGN.LEFT):
    tf = shape_or_tf.text_frame if hasattr(shape_or_tf, "text_frame") else shape_or_tf
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.name = "Calibri"
    r.font.color.rgb = color
    return tf


def add_logo(slide, logo_bytes):
    h = 0.5
    try:
        w_px, h_px = Image.open(io.BytesIO(logo_bytes)).size
        w = h * w_px / h_px
        slide.shapes.add_picture(io.BytesIO(logo_bytes), Inches(SLIDE_W - MARGIN + 0.1 - w), Inches(0.08),
                                 Inches(w), Inches(h))
    except Exception:
        tb = slide.shapes.add_textbox(Inches(BAR_W + 0.2), Inches(0.08), Inches(SLIDE_W - BAR_W - 0.4), Inches(h))
        tb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        _text(tb, "Channelplay", 22, BLUE, bold=True, align=PP_ALIGN.RIGHT)


def add_header(slide, title, logo_bytes):
    bar = _rect(slide, 0, 0, BAR_W, BAR_H, BLUE)
    tf = bar.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.25)
    size = 20 if len(title) <= 58 else 18 if len(title) <= 70 else 15
    _text(tf, title, size, RGBColor(255, 255, 255))
    add_logo(slide, logo_bytes)


def add_welcome(prs, title, subtitle, logo_bytes):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_logo(slide, logo_bytes)
    band = _rect(slide, 0, 2.7, SLIDE_W, 1.5, BLUE)
    band.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    band.text_frame.word_wrap = True
    _text(band, title, 36, RGBColor(255, 255, 255), bold=True, align=PP_ALIGN.CENTER)
    if subtitle:
        tb = slide.shapes.add_textbox(Inches(1), Inches(4.4), Inches(SLIDE_W - 2), Inches(0.7))
        tb.text_frame.word_wrap = True
        _text(tb, subtitle, 20, GREY, align=PP_ALIGN.CENTER)


def add_thanks(prs, logo_bytes):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_logo(slide, logo_bytes)
    band = _rect(slide, 0, 2.7, SLIDE_W, 1.5, BLUE)
    band.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    _text(band, "Thank You", 40, RGBColor(255, 255, 255), bold=True, align=PP_ALIGN.CENTER)


def add_image_box(slide, img, box):
    """Place one prepared image (or a 'no image' placeholder) fitted inside box."""
    bx, by, bw, bh = box
    if img is None:
        ph = _rect(slide, bx, by, bw, bh, LIGHT)
        ph.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        _text(ph, "Image not available", 12, GREY, align=PP_ALIGN.CENTER)
        return
    data, (iw, ih) = img
    scale = min(bw / iw, bh / ih)
    w, h = iw * scale, ih * scale
    pic = slide.shapes.add_picture(io.BytesIO(data), Inches(bx + (bw - w) / 2), Inches(by + (bh - h) / 2),
                                   Inches(w), Inches(h))
    pic.line.color.rgb = GREY
    pic.line.width = Pt(0.75)


def add_outlet_slide(prs, entry, logo_bytes, images, active, show_titles=False):
    """entry: {title, front, before, after} (url lists); images: url -> prepared image or None.
    active: the (key, label) columns that are switched on; they share the slide width equally."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide, entry["title"], logo_bytes)

    n = len(active)
    col_w = (SLIDE_W - 2 * MARGIN - (n - 1) * COL_GAP) / n
    if show_titles:
        label_y, label_h = BAR_H + 0.2, 0.35
        area_y = label_y + label_h + 0.12
    else:
        area_y = BAR_H + 0.3
    area_h = SLIDE_H - area_y - 0.3
    for i, (key, label) in enumerate(active):
        x = MARGIN + i * (col_w + COL_GAP)
        if show_titles:
            tb = slide.shapes.add_textbox(Inches(x), Inches(label_y), Inches(col_w), Inches(label_h))
            tb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            _text(tb, label, 14, BLUE, bold=True, align=PP_ALIGN.CENTER)
            _rect(slide, x, label_y + label_h, col_w, 0.03, BLUE)

        urls = entry[key]
        if not urls:
            add_image_box(slide, None, (x, area_y, col_w, area_h))
            continue
        for url, box in zip(urls, boxes_for(len(urls), x, area_y, col_w, area_h)):
            add_image_box(slide, images.get(url), box)


def build_deck(entries, images, logo_bytes, welcome_sub="", active=COLUMNS, show_titles=False):
    """Welcome slide per campaign (name from the Excel) + its outlet slides, then a closing Thank You slide."""
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(SLIDE_W), Inches(SLIDE_H)
    groups = {}
    for e in entries:
        groups.setdefault(e["campaign"], []).append(e)
    for campaign, items in groups.items():
        if campaign:
            add_welcome(prs, campaign, welcome_sub, logo_bytes)
        for e in items:
            add_outlet_slide(prs, e, logo_bytes, images, active, show_titles)
    add_thanks(prs, logo_bytes)
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def build_entries(ws, headers, cols, rows=None, active_keys=None):
    """One entry per Excel row (identical rows skipped).

    cols maps: front/before/after (image columns), code, client, name, city, state, campaign.
    Header text = "<client code> – <store name> – <city>, <state>"; campaign = welcome-slide title.
    """
    idx = {h: i + 1 for i, h in enumerate(headers) if h}
    entries, seen = [], set()
    for row in range(2, ws.max_row + 1):
        if rows is not None and row not in rows:
            continue

        def text(key):
            h = cols.get(key)
            if not h or h not in idx:
                return ""
            v = ws.cell(row=row, column=idx[h]).value
            return "" if v is None else str(v).strip()

        urls = {k: (cell_urls(ws.cell(row=row, column=idx[cols[k]])) if active_keys is None or k in active_keys else [])
                for k, _ in COLUMNS}
        place = ", ".join(p for p in (text("city"), text("state")) if p)
        title = " – ".join(p for p in (text("client"), text("name"), place) if p) or f"Row {row}"
        code = text("code") or f"row{row}"
        campaign = text("campaign")
        key = (code, title, campaign, tuple(urls["front"]), tuple(urls["before"]), tuple(urls["after"]))
        if key in seen:
            continue
        seen.add(key)
        entries.append({"code": code, "name": text("name"), "title": title, "campaign": campaign, **urls})
    return entries


def download_all(urls, workers=8):
    """url -> prepared image (or None). Returns (images, failures)."""
    progress = st.progress(0.0, text=f"Downloading 0 of {len(urls)} images...")
    images, failures = {}, []
    jobs = [{"url": u, "stem": "x"} for u in urls]
    with ThreadPoolExecutor(max_workers=workers) as pool:
        for done, res in enumerate(pool.map(fetch, jobs), start=1):
            prepared = prep_image(res["data"]) if not res["error"] else None
            images[res["url"]] = prepared
            if prepared is None:
                failures.append({"Link": res["url"], "Reason": res["error"] or "not a readable image"})
            progress.progress(done / len(urls), text=f"Downloading {done} of {len(urls)} images...")
    progress.empty()
    return images, failures


# --------------------------------------------------------------------------
# UI
# --------------------------------------------------------------------------
def pick(headers, *needles, default=0):
    for i, h in enumerate(headers):
        if h and all(n in h.lower() for n in needles):
            return i
    return default


def main():
    st.set_page_config(page_title="Images to ZIP & PPT Maker", page_icon="🗜️", layout="centered")

    tab_excel, tab_links, tab_ppt = st.tabs(["📄 Upload Excel", "🔗 Paste links", "📊 Make PPT"])

    # ---------------- Excel -> ZIP ----------------
    with tab_excel:
        st.title("Images to ZIP")
        st.write(
            "Upload your Excel. Every image column is downloaded, including the "
            "“N photos — open” gallery cells. Files are named **StoreCode_ColumnName.jpg**."
        )
        up = st.file_uploader("Excel file", type=["xlsx", "xlsm"], key="up_zip")
        if up:
            file_bytes = up.getvalue()
            sheets = load_workbook(io.BytesIO(file_bytes), read_only=True).sheetnames
            sheet = st.selectbox("Sheet", sheets, key="sh_zip") if len(sheets) > 1 else sheets[0]
            wb, ws, headers = read_workbook(file_bytes, sheet)
            named = [h for h in headers if h]

            code_col = st.selectbox("Outlet code column", named, index=pick(named, "store code"), key="code_zip")
            auto_img = [h for h in named if "image" in h.lower()]
            image_cols = st.multiselect("Image columns", named, default=auto_img)

            rows = date_filter_ui(ws, headers, "zip")
            jobs = build_jobs(ws, headers, code_col, image_cols, rows) if image_cols else []
            st.caption(f"{len(jobs)} image{'s' if len(jobs) != 1 else ''} found in {ws.max_row - 1 if rows is None else len(rows)} rows")
            if jobs:
                with st.expander("Preview file names"):
                    st.dataframe([{"File": j["stem"] + ".jpg", "Link": j["url"]} for j in jobs], hide_index=True)

            zip_name_x = st.text_input("ZIP file name", value="images", key="zip_x")
            if st.button("Prepare ZIP", type="primary", disabled=not jobs, key="go_x"):
                st.session_state.pop("zip_result", None)
                run_jobs(jobs, zip_name_x)

    # ---------------- Paste links -> ZIP ----------------
    with tab_links:
        st.title("Images to ZIP")
        st.write("Paste image links, one per line. Get them all in a single ZIP file.")
        raw = st.text_area("Image URLs", height=220,
                           placeholder="https://example.com/photo1.jpg\nhttps://example.com/photo2.png")
        links = parse_links(raw)

        col1, col2 = st.columns([3, 1], vertical_alignment="bottom")
        zip_name_l = col1.text_input("ZIP file name", value="images", key="zip_l")
        col2.caption(f"{len(links)} link{'s' if len(links) != 1 else ''}")

        if st.button("Prepare ZIP", type="primary", disabled=not links, key="go_l"):
            st.session_state.pop("zip_result", None)
            jobs = []
            for i, u in enumerate(links):
                base = re.sub(r"[^\w.\-]+", "_", unquote(Path(urlparse(u).path).name) or "image")
                jobs.append({"url": u, "stem": f"{i + 1:03d}_{Path(base).stem}", "code": "", "column": ""})
            run_jobs(jobs, zip_name_l)

    # ---------------- ZIP result (shared by the two tabs above) ----------------
    result = st.session_state.get("zip_result")
    if result:
        st.divider()
        if result["bytes"]:
            msg = f"{result['good']} of {result['total']} images are ready."
            (st.success if result["good"] == result["total"] else st.warning)(msg)
            st.download_button("Download ZIP", data=result["bytes"], file_name=result["name"],
                               mime="application/zip", type="primary")
        else:
            st.error("No images could be downloaded. Check that the links point directly to image files.")
        st.dataframe(result["report"], width="stretch", hide_index=True)

    # ---------------- Excel -> PPT ----------------
    with tab_ppt:
        st.title("✨ Outlet Deck Maker")
        st.subheader("From Excel to a client-ready PPT in one click")
        st.write("Upload your Excel to get a PowerPoint with **one slide per outlet**: "
                 "store front, then before deployment, then after deployment.")
        up = st.file_uploader("Excel file", type=["xlsx", "xlsm"], key="up_ppt")
        if up:
            file_bytes = up.getvalue()
            sheets = load_workbook(io.BytesIO(file_bytes), read_only=True).sheetnames
            sheet = st.selectbox("Sheet", sheets, key="sh_ppt") if len(sheets) > 1 else sheets[0]
            wb, ws, headers = read_workbook(file_bytes, sheet)
            named = [h for h in headers if h]

            st.markdown("**Images to include** (switch off what you don't want on the slides)")
            c1, c2, c3 = st.columns(3)
            front_col = c1.selectbox("Store front column", named, index=pick(named, "front"))
            on_front = c1.toggle("Store front", value=True)
            before_col = c2.selectbox("Before column", named, index=pick(named, "before"))
            on_before = c2.toggle("Before deployment", value=True)
            after_col = c3.selectbox("After column", named, index=pick(named, "after"))
            on_after = c3.toggle("After deployment", value=True)
            active = [(k, l) for (k, l), on in zip(COLUMNS, (on_front, on_before, on_after)) if on]
            if not active:
                st.warning("Switch on at least one image type.")
            show_titles = st.checkbox("Show column titles on slides (Store Front / Before / After)", value=False)

            rows = date_filter_ui(ws, headers, "ppt")

            code_col = st.selectbox("Outlet code column (used for file names)", named,
                                    index=pick(named, "store code"), key="code_ppt")

            st.markdown("**Blue header text:** client code – store name – city, state")
            h1, h2, h3, h4 = st.columns(4)
            NONE = "(none)"
            opts = [NONE] + named
            def opt_idx(*needles):
                """Index into opts for the first header containing all needles, else 0 = (none)."""
                i = pick(named, *needles, default=-1)
                return i + 1 if i >= 0 else 0

            client_col = h1.selectbox("Client code", opts, index=opt_idx("client", "code"))
            name_col = h2.selectbox("Store name", opts, index=opt_idx("store name"))
            city_col = h3.selectbox("City", opts, index=opt_idx("city"))
            state_col = h4.selectbox("State", opts, index=opt_idx("state"))

            st.markdown("**Welcome slide**")
            camp_col = st.selectbox("Campaign column (its value becomes the welcome-slide title)", opts,
                                    index=opt_idx("campaign", "name"))
            w_sub = st.text_input("Subtitle (optional)", value="")
            logo_up = st.file_uploader("Logo (optional – a sharper Channelplay logo, PNG/JPG)",
                                       type=["png", "jpg", "jpeg"])
            logo_bytes = logo_up.getvalue() if logo_up else base64.b64decode(DEFAULT_LOGO_B64)

            mode = st.radio("Output", ["One PPT with all outlets", "One PPT per outlet (ZIP)"], horizontal=True)

            cols = {"front": front_col, "before": before_col, "after": after_col, "code": code_col,
                    "client": None if client_col == NONE else client_col,
                    "name": None if name_col == NONE else name_col,
                    "city": None if city_col == NONE else city_col,
                    "state": None if state_col == NONE else state_col,
                    "campaign": None if camp_col == NONE else camp_col}
            entries = build_entries(ws, headers, cols, rows, {k for k, _ in active}) if active else []
            if entries:
                st.caption("Header preview: " + entries[0]["title"]
                           + (f"   |   Welcome slide: {entries[0]['campaign']}" if entries[0]["campaign"] else ""))
            all_urls = list(dict.fromkeys(u for e in entries for k, _ in active for u in e[k]))
            st.caption(f"{len(entries)} outlet slide{'s' if len(entries) != 1 else ''}, {len(all_urls)} images")

            if st.button("Create PPT", type="primary", disabled=not entries, key="go_ppt"):
                st.session_state.pop("ppt_result", None)
                images, failures = download_all(all_urls)
                with st.spinner("Building PowerPoint..."):
                    if mode.startswith("One PPT with"):
                        data = build_deck(entries, images, logo_bytes, w_sub, active, show_titles)
                        name, mime = "outlet_report.pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    else:
                        groups = {}
                        for e in entries:
                            groups.setdefault(e["code"], []).append(e)
                        buf = io.BytesIO()
                        with zipfile.ZipFile(buf, "w", zipfile.ZIP_STORED) as zf:
                            for code, items in groups.items():
                                zf.writestr(f"{clean(code)}_{clean(items[0]['name'])}.pptx".replace("_.pptx", ".pptx"),
                                            build_deck(items, images, logo_bytes, w_sub, active, show_titles))
                        data, name, mime = buf.getvalue(), "outlet_ppts.zip", "application/zip"
                st.session_state["ppt_result"] = {"bytes": data, "name": name, "mime": mime,
                                                  "failures": failures, "total": len(all_urls)}

        pres = st.session_state.get("ppt_result")
        if pres:
            st.divider()
            ok = pres["total"] - len(pres["failures"])
            (st.success if not pres["failures"] else st.warning)(
                f"PPT ready. {ok} of {pres['total']} images placed."
                + ("" if not pres["failures"] else " Missing ones show “Image not available”."))
            st.download_button("Download PPT", data=pres["bytes"], file_name=pres["name"],
                               mime=pres["mime"], type="primary")
            if pres["failures"]:
                st.dataframe(pres["failures"], width="stretch", hide_index=True)


if __name__ == "__main__":
    main()
