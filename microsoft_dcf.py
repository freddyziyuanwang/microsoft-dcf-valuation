import yfinance as yf
import pandas as pd

# 下载微软的财务数据（Download Microsoft financial data）
ticker = 'MSFT'
msft = yf.Ticker(ticker)

# 获取财务报表数据（Grab financial statement data）
income_statement = msft.financials.T  # 利润表 (Income Statement)
balance_sheet = msft.balance_sheet.T  # 资产负债表 (Balance Sheet)
cash_flow = msft.cashflow.T           # 现金流量表 (Cash Flow Statement)

# 打印现金流量表的列名（Inspect column names in Cash Flow Statement）
print("现金流量表的列名 (Columns in Cash Flow Statement):")
print(cash_flow.columns)

# 动态匹配列名，确保列名正确（Dynamically match column names to avoid errors due to name changes）
operating_cash_flow_col = "Operating Cash Flow"
capital_expenditures_col = "Capital Expenditure"

# 填充缺失值为 0（Fill missing values with 0）
income_statement = income_statement.fillna(0)
cash_flow = cash_flow.fillna(0)

# 确保年份索引一致（Align index by fiscal years）
common_index = income_statement.index.intersection(cash_flow.index)

# 提取关键财务指标（Extract key financial metrics）
data = {
    'Year': common_index.year,  # 财务年份（Fiscal Year）
    'Revenue': income_statement.loc[common_index, 'Total Revenue'],  # 总收入 (Total Revenue)
    'Operating Income': income_statement.loc[common_index, 'Operating Income'],  # 运营利润 (EBIT, close to EBITDA)
    'Net Income': income_statement.loc[common_index, 'Net Income'],  # 净利润 (Net Income)
    'Operating Cash Flow': cash_flow.loc[common_index, operating_cash_flow_col],  # 经营现金流 (Operating Cash Flow)
    'Capital Expenditures': cash_flow.loc[common_index, capital_expenditures_col]  # 资本支出 (Capital Expenditures)
}

# 创建一个 DataFrame（Create a DataFrame）
financial_data = pd.DataFrame(data)

# 计算自由现金流（Calculate Free Cash Flow, FCF）
financial_data['Free Cash Flow'] = financial_data['Operating Cash Flow'] + financial_data['Capital Expenditures']

# 打印整理后的财务数据（Print organized financial data）
print("\n微软财务数据整理 (Organized Microsoft Financial Data):")
print(financial_data)

# 保存数据到 Excel 文件（Save data to an Excel file）
financial_data.to_excel("microsoft_financials.xlsx", index=False)
print("数据已保存到 microsoft_financials.xlsx 文件 (Data has been saved to microsoft_financials.xlsx)！")
import pandas as pd

# 从 Excel 文件加载数据（Load data from Excel file）
financial_data = pd.read_excel("microsoft_financials.xlsx")

# 计算年均增长率（Calculate Compound Annual Growth Rate, CAGR）
# CAGR 公式: (最后一年 FCF / 第一年 FCF)^(1 / (年数 - 1)) - 1
fcf_cagr = (financial_data['Free Cash Flow'].iloc[-1] / financial_data['Free Cash Flow'].iloc[0]) ** (1 / (len(financial_data) - 1)) - 1
print(f"自由现金流年均增长率 (FCF CAGR): {fcf_cagr:.2%}")

# 预测未来 5 年的自由现金流（Forecast Free Cash Flow for next 5 years）
future_years = [2025, 2026, 2027, 2028, 2029]
future_fcf = []
last_fcf = financial_data['Free Cash Flow'].iloc[-1]

# 使用 CAGR 预测未来自由现金流（Use CAGR to predict future FCF）
for year in future_years:
    next_fcf = last_fcf * (1 + fcf_cagr)
    future_fcf.append(next_fcf)
    last_fcf = next_fcf

# 构建预测数据表（Create a forecast DataFrame）
forecast_data = pd.DataFrame({
    'Year': future_years,
    'Predicted FCF': future_fcf
})

# 将预测结果与原始数据结合（Combine forecast with original data）
financial_data['Predicted FCF'] = [None] * len(financial_data)  # 占位列
forecast_combined = pd.concat([financial_data[['Year', 'Free Cash Flow']], forecast_data], ignore_index=True)

# 输出预测结果（Print the forecast result）
print("\n未来自由现金流预测 (Future Free Cash Flow Forecast):")
print(forecast_combined)

# 将结果保存到新的 Excel 文件（Save the forecast to a new Excel file）
forecast_combined.to_excel("microsoft_fcf_forecast.xlsx", index=False)
print("预测结果已保存到 microsoft_fcf_forecast.xlsx 文件 (Forecast has been saved to microsoft_fcf_forecast.xlsx)！")
import numpy as np

# 假设输入参数（Assumed Inputs）
risk_free_rate = 0.04  # 无风险利率，假设为 4% (Risk-Free Rate, assumed to be 4%)
market_risk_premium = 0.06  # 市场风险溢价，假设为 6% (Market Risk Premium, assumed to be 6%)
beta = 0.9  # 微软的 Beta 值，反映系统性风险 (Microsoft's Beta, reflecting systematic risk)
cost_of_equity = risk_free_rate + beta * market_risk_premium  # 股权成本公式 (Cost of Equity formula)

# 债务相关参数（Debt-Related Inputs）
cost_of_debt = 0.03  # 假设债务成本为 3% (Cost of Debt, assumed to be 3%)
tax_rate = 0.21  # 假设企业税率为 21% (Corporate Tax Rate, assumed to be 21%)

# 资本结构（Capital Structure）
equity_value = 2.5e12  # 微软市值（假设为 2.5 万亿美元）(Microsoft's market cap, assumed to be $2.5 trillion)
debt_value = 0.05e12  # 假设总债务为 500 亿美元 (Total debt, assumed to be $50 billion)
total_value = equity_value + debt_value  # 总资本价值 (Total Value of Capital)

# 计算 WACC（Calculate WACC）
equity_weight = equity_value / total_value  # 股权权重 (Equity Weight)
debt_weight = debt_value / total_value  # 债务权重 (Debt Weight)

wacc = equity_weight * cost_of_equity + debt_weight * cost_of_debt * (1 - tax_rate)
print(f"加权平均资本成本 (WACC): {wacc:.2%}")
# 假设终值增长率（Assumed Terminal Growth Rate）
terminal_growth_rate = 0.025  # 假设长期增长率为 2.5% (Assumed Terminal Growth Rate)

# 获取最后一年的自由现金流（Get the last year's Free Cash Flow）
last_year_fcf = forecast_combined['Predicted FCF'].iloc[-1]

# 计算终值（Calculate Terminal Value）
terminal_value = last_year_fcf * (1 + terminal_growth_rate) / (wacc - terminal_growth_rate)
print(f"终值 (Terminal Value): {terminal_value:.2e}")

# 将未来现金流和终值贴现到当前时点（Discount Future Cash Flows and Terminal Value to Present Value）
discounted_fcf = []  # 存储每年的贴现现金流 (Store discounted cash flows)

for t, fcf in enumerate(forecast_combined['Predicted FCF'].dropna(), start=1):
    pv = fcf / (1 + wacc) ** t
    discounted_fcf.append(pv)

# 折现终值（Discount Terminal Value）
terminal_value_pv = terminal_value / (1 + wacc) ** len(future_years)
print(f"折现终值 (Discounted Terminal Value): {terminal_value_pv:.2e}")

# 合计估值（Calculate Total Valuation）
total_valuation = sum(discounted_fcf) + terminal_value_pv
print(f"企业估值 (Enterprise Valuation): {total_valuation:.2e}")

# 当前市场价格（Current Market Valuation）
current_market_valuation = 2.5e12  # 假设微软当前市值为 2.5 万亿美元 (Assume Microsoft market cap is $2.5 trillion)

# 比较估值与市场价格（Compare DCF valuation to market valuation）
if total_valuation > current_market_valuation:
    print(f"DCF 模型估值 ({total_valuation:.2e}) 高于市场价格 ({current_market_valuation:.2e})，可能被低估 (Potentially Undervalued)")
elif total_valuation < current_market_valuation:
    print(f"DCF 模型估值 ({total_valuation:.2e}) 低于市场价格 ({current_market_valuation:.2e})，可能被高估 (Potentially Overvalued)")
else:
    print(f"DCF 模型估值与市场价格接近，可能为合理估值 (Fairly Valued)")
import matplotlib.pyplot as plt

# 图表 1：未来自由现金流预测（Future Free Cash Flow Forecast）
plt.figure(figsize=(10, 6))
plt.plot(forecast_combined['Year'], forecast_combined['Free Cash Flow'], label="历史 FCF (Historical FCF)", marker='o')
plt.plot(forecast_combined['Year'], forecast_combined['Predicted FCF'], label="预测 FCF (Forecasted FCF)", linestyle='--', marker='x')
plt.title("微软自由现金流预测 (Microsoft Free Cash Flow Forecast)")
plt.xlabel("年份 (Year)")
plt.ylabel("自由现金流 (Free Cash Flow, USD)")
plt.legend()
plt.grid(True)
plt.savefig("microsoft_fcf_forecast.png")
plt.show()

# 图表 2：估值对比（Valuation Comparison）
plt.figure(figsize=(6, 4))
labels = ['DCF 估值 (DCF Valuation)', '市场价格 (Market Price)']
values = [total_valuation, current_market_valuation]
plt.bar(labels, values, color=['blue', 'orange'])
plt.title("估值结果对比 (Valuation Comparison)")
plt.ylabel("估值 (Valuation, USD)")
plt.savefig("microsoft_valuation_comparison.png")
plt.show()
# 生成估值报告（Generate Valuation Report）
with open("microsoft_valuation_report.txt", "w") as file:
    file.write("微软 DCF 模型估值报告 (Microsoft DCF Valuation Report)\n")
    file.write("=" * 50 + "\n")
    file.write(f"DCF 模型估值: {total_valuation:.2e} USD\n")
    file.write(f"当前市场价格: {current_market_valuation:.2e} USD\n")
    if total_valuation > current_market_valuation:
        file.write("结论: 可能被低估 (Potentially Undervalued)\n")
    elif total_valuation < current_market_valuation:
        file.write("结论: 可能被高估 (Potentially Overvalued)\n")
    else:
        file.write("结论: 可能为合理估值 (Fairly Valued)\n")
    file.write("=" * 50 + "\n")
    file.write("预测未来自由现金流 (Predicted Future Free Cash Flow):\n")
    file.write(forecast_combined.to_string(index=False))
print("估值报告已生成 (Valuation Report Generated): microsoft_valuation_report.txt")
