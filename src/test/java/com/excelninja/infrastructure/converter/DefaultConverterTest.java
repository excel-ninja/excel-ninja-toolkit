package com.excelninja.infrastructure.converter;

import com.excelninja.domain.exception.DocumentConversionException;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Date;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

@DisplayName("BigDecimal과 LocalDateTime 지원 테스트")
class DefaultConverterTest {

    private final DefaultConverter converter = new DefaultConverter();

    @Test
    @DisplayName("BigDecimal 변환 테스트")
    void convertToBigDecimal() {
        assertThat(converter.convert("123.456", BigDecimal.class)).isEqualTo(new BigDecimal("123.456"));
        assertThat(converter.convert("999999999999999999.999999999", BigDecimal.class)).isEqualTo(new BigDecimal("999999999999999999.999999999"));
        assertThat(converter.convert(123.45, BigDecimal.class)).isEqualTo(BigDecimal.valueOf(123.45));
        assertThat(converter.convert(100, BigDecimal.class)).isEqualTo(BigDecimal.valueOf(100));
        assertThat(converter.convert(123L, BigDecimal.class)).isEqualTo(BigDecimal.valueOf(123));

        BigDecimal originalBigDecimal = new BigDecimal("999.99");
        assertThat(converter.convert(originalBigDecimal, BigDecimal.class)).isSameAs(originalBigDecimal);

        assertThatThrownBy(() -> converter.convert("invalid_number", BigDecimal.class))
                .isInstanceOf(DocumentConversionException.class)
                .hasMessageContaining("Cannot parse 'invalid_number' as BigDecimal");
    }

    @Test
    @DisplayName("LocalDate 변환 테스트")
    void convertToLocalDate() {
        LocalDate lastChristmas = LocalDate.of(2024, 12, 25);
        assertThat(converter.convert("2024-12-25", LocalDate.class)).isEqualTo(lastChristmas);
        assertThat(converter.convert("25/12/2024", LocalDate.class)).isEqualTo(lastChristmas);
        assertThat(converter.convert("12/25/2024", LocalDate.class)).isEqualTo(lastChristmas);
        assertThat(converter.convert("2024/12/25", LocalDate.class)).isEqualTo(lastChristmas);
        assertThat(converter.convert("25-12-2024", LocalDate.class)).isEqualTo(lastChristmas);

        Date date = new Date(124, 11, 25);
        LocalDate convertedDate = (LocalDate) converter.convert(date, LocalDate.class);
        assertThat(convertedDate.getYear()).isEqualTo(2024);
        assertThat(convertedDate.getMonthValue()).isEqualTo(12);
        assertThat(convertedDate.getDayOfMonth()).isEqualTo(25);

        LocalDateTime dateTime = LocalDateTime.of(2024, 12, 25, 14, 30, 0);
        assertThat(converter.convert(dateTime, LocalDate.class)).isEqualTo(lastChristmas);

        assertThatThrownBy(() -> converter.convert("invalid_date", LocalDate.class))
                .isInstanceOf(DocumentConversionException.class)
                .hasMessageContaining("Cannot parse date string 'invalid_date'");
    }

    @Test
    @DisplayName("LocalDateTime 변환 테스트")
    void convertToLocalDateTime() {
        LocalDateTime lastChristmasTwoThirty = LocalDateTime.of(2024, 12, 25, 14, 30, 0);
        assertThat(converter.convert("2024-12-25T14:30:00", LocalDateTime.class)).isEqualTo(lastChristmasTwoThirty);
        assertThat(converter.convert("2024-12-25 14:30:00", LocalDateTime.class)).isEqualTo(lastChristmasTwoThirty);
        assertThat(converter.convert("2024-12-25 14:30", LocalDateTime.class)).isEqualTo(lastChristmasTwoThirty);
        assertThat(converter.convert("25/12/2024 14:30:00", LocalDateTime.class)).isEqualTo(lastChristmasTwoThirty);
        assertThat(converter.convert("12/25/2024 14:30:00", LocalDateTime.class)).isEqualTo(lastChristmasTwoThirty);

        Date date = new Date(124, 11, 25, 14, 30, 0);
        LocalDateTime convertedDateTime = (LocalDateTime) converter.convert(date, LocalDateTime.class);
        assertThat(convertedDateTime.getYear()).isEqualTo(2024);
        assertThat(convertedDateTime.getMonthValue()).isEqualTo(12);
        assertThat(convertedDateTime.getDayOfMonth()).isEqualTo(25);
        assertThat(convertedDateTime.getHour()).isEqualTo(14);
        assertThat(convertedDateTime.getMinute()).isEqualTo(30);

        LocalDate date2 = LocalDate.of(2024, 12, 25);
        assertThat(converter.convert(date2, LocalDateTime.class)).isEqualTo(LocalDateTime.of(2024, 12, 25, 0, 0, 0));

        assertThatThrownBy(() -> converter.convert("invalid_datetime", LocalDateTime.class))
                .isInstanceOf(DocumentConversionException.class)
                .hasMessageContaining("Cannot parse datetime string 'invalid_datetime'");
    }
}