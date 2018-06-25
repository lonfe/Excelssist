package slf.excel;

import java.lang.annotation.*;

@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface Cell {
    String value() default "";
    int num() default 0;
}
