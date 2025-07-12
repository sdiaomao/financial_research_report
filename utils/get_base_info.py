import efinance as ef


def get_base_info(stock_code: str = "000001"):
    """获取股票基础信息"""

    return ef.stock.get_base_info(stock_code)



if __name__ == "__main__":
    data = get_base_info()
    print(data)
