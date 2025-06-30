package com.excelninja.application.port;


public interface ConverterPort {
    Object convert(Object rawValue, Class<?> targetType);
}
