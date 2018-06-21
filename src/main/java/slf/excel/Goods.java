package slf.excel;

import java.math.BigDecimal;
import java.util.Date;

/**
 * @auther shenlf
 * @create 2018/6/22 1:27
 */
public class Goods {
    @Sign(num=1)
    private String name;
    @Sign(num=2)
    private BigDecimal price;
    @Sign(num=3)
    private Integer count;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public BigDecimal getPrice() {
        return price;
    }

    public void setPrice(BigDecimal price) {
        this.price = price;
    }

    public Integer getCount() {
        return count;
    }

    public void setCount(Integer count) {
        this.count = count;
    }

}
