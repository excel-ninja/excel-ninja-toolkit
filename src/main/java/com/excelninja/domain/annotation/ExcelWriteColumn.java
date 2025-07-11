package com.excelninja.domain.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelWriteColumn {
    String headerName();           // 헤더명
    int order() default Integer.MAX_VALUE;
}
