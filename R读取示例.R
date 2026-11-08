# ==============================================================================
# 在 R 里读取 FPF数据处理.exe 导出的结果
# 列名、列类型(Date / factor / integer / numeric / character)与原 R 脚本
# clean_production_data() 的输出一致, 后续分析/作图代码可以直接沿用。
# rds 用 readRDS() 读取, 无需装包。
# ==============================================================================

dat <- readRDS("D:/数据/综合选择查询_清洗.rds")

# 转成 data.table (可选)
data.table::setDT(dat)

str(dat)
levels(dat$猪场)      # 因子水平 = farm_order, 作图顺序固定
class(dat$配种日期)   # "Date"

# csv 导出时:
# dat <- data.table::fread("D:/数据/综合选择查询_清洗.csv", encoding = "UTF-8")
