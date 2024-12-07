from openpyxl.styles import Border, Side

thin = Side(border_style="thin", color="000000")  # 细实线，黑色
round_thin_border = Border(left=thin, right=thin, top=thin, bottom=thin)