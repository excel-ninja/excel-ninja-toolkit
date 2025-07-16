package com.excelninja.infrastructure.util;

import com.excelninja.domain.exception.DocumentConversionException;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.lang.reflect.Field;

import static org.assertj.core.api.Assertions.*;

class ReflectionUtilsTest {

    @BeforeEach
    void setUp() {
        ReflectionUtils.clearMethodCache();
    }

    @AfterEach
    void tearDown() {
        ReflectionUtils.clearMethodCache();
    }

    @Test
    @DisplayName("Setter 메서드를 통한 필드 값 설정")
    void setFieldValueUsingSetterMethod() {
        // Given
        TestDtoWithSetter dto = new TestDtoWithSetter();
        Field field = getField(TestDtoWithSetter.class, "name");

        // When
        ReflectionUtils.setFieldValue(dto, field, "Hyunsoo");

        // Then
        assertThat(dto.getName()).isEqualTo("Hyunsoo");
        assertThat(dto.setterCalled).isTrue();
    }

    @Test
    @DisplayName("Getter 메서드를 통한 필드 값 조회")
    void getFieldValueUsingGetterMethod() {
        // Given
        TestDtoWithGetter dto = new TestDtoWithGetter();
        dto.setName("Jane");
        Field field = getField(TestDtoWithGetter.class, "name");

        // When
        Object value = ReflectionUtils.getFieldValue(dto, field);

        // Then
        assertThat(value).isEqualTo("Jane");
        assertThat(dto.getterCalled).isTrue();
    }

    @Test
    @DisplayName("Boolean 필드의 is 메서드 사용")
    void getBooleanFieldUsingIsMethod() {
        // Given
        TestDtoWithBooleanGetter dto = new TestDtoWithBooleanGetter();
        dto.setActive(true);
        Field field = getField(TestDtoWithBooleanGetter.class, "active");

        // When
        Object value = ReflectionUtils.getFieldValue(dto, field);

        // Then
        assertThat(value).isEqualTo(true);
        assertThat(dto.isMethodCalled).isTrue();
    }

    @Test
    @DisplayName("Setter가 없을 때 직접 필드 접근")
    void setFieldValueDirectlyWhenNoSetter() {
        // Given
        TestDtoWithoutSetter dto = new TestDtoWithoutSetter();
        Field field = getField(TestDtoWithoutSetter.class, "name");

        // When
        ReflectionUtils.setFieldValue(dto, field, "Direct");

        // Then
        assertThat(dto.name).isEqualTo("Direct");
    }

    @Test
    @DisplayName("Getter가 없을 때 직접 필드 접근")
    void getFieldValueDirectlyWhenNoGetter() {
        // Given
        TestDtoWithoutGetter dto = new TestDtoWithoutGetter();
        dto.name = "Direct";
        Field field = getField(TestDtoWithoutGetter.class, "name");

        // When
        Object value = ReflectionUtils.getFieldValue(dto, field);

        // Then
        assertThat(value).isEqualTo("Direct");
    }

    @Test
    @DisplayName("Final 필드 설정 시 예외 발생")
    void setFinalFieldThrowsException() {
        // Given
        TestDtoWithFinalField dto = new TestDtoWithFinalField();
        Field field = getField(TestDtoWithFinalField.class, "finalValue");

        // When & Then
        assertThatThrownBy(() -> ReflectionUtils.setFieldValue(dto, field, "NewValue"))
                .isInstanceOf(DocumentConversionException.class)
                .hasMessageContaining("Unexpected error setting field: finalValue");
    }


    @Test
    @DisplayName("타입 불일치 시 예외 발생")
    void setFieldValueWithWrongTypeThrowsException() {
        // Given
        TestDtoWithSetter dto = new TestDtoWithSetter();
        Field field = getField(TestDtoWithSetter.class, "age");

        // When & Then
        assertThatThrownBy(() -> ReflectionUtils.setFieldValue(dto, field, "NotAnInteger"))
                .isInstanceOf(DocumentConversionException.class)
                .hasMessageContaining("Unexpected error setting field");
    }

    @Test
    @DisplayName("Null 인스턴스 시 예외 발생")
    void setFieldValueWithNullInstanceThrowsException() {
        // Given
        Field field = getField(TestDtoWithSetter.class, "name");

        // When & Then
        assertThatThrownBy(() -> ReflectionUtils.setFieldValue(null, field, "Value"))
                .isInstanceOf(DocumentConversionException.class)
                .hasMessageContaining("Instance cannot be null");
    }

    @Test
    @DisplayName("Null 필드 시 예외 발생")
    void setFieldValueWithNullFieldThrowsException() {
        // Given
        TestDtoWithSetter dto = new TestDtoWithSetter();

        // When & Then
        assertThatThrownBy(() -> ReflectionUtils.setFieldValue(dto, null, "Value"))
                .isInstanceOf(DocumentConversionException.class)
                .hasMessageContaining("Field cannot be null");
    }

    @Test
    @DisplayName("메서드 캐싱 동작 확인")
    void methodCachingWorks() {
        // Given
        TestDtoWithSetter dto1 = new TestDtoWithSetter();
        TestDtoWithSetter dto2 = new TestDtoWithSetter();
        Field field = getField(TestDtoWithSetter.class, "name");

        // When
        ReflectionUtils.setFieldValue(dto1, field, "First");
        int cacheCountAfterFirst = ReflectionUtils.getCachedMethodCount();

        ReflectionUtils.setFieldValue(dto2, field, "Second");
        int cacheCountAfterSecond = ReflectionUtils.getCachedMethodCount();

        // Then
        assertThat(dto1.getName()).isEqualTo("First");
        assertThat(dto2.getName()).isEqualTo("Second");
        assertThat(cacheCountAfterFirst).isEqualTo(1);
        // cached value
        assertThat(cacheCountAfterSecond).isEqualTo(1);
    }

    @Test
    @DisplayName("기본 타입과 래퍼 타입 간 호환성")
    void primitiveAndWrapperTypeCompatibility() {
        // Given
        TestDtoWithPrimitives dto = new TestDtoWithPrimitives();
        Field intField = getField(TestDtoWithPrimitives.class, "primitiveInt");
        Field IntegerField = getField(TestDtoWithPrimitives.class, "wrapperInteger");

        // When & Then
        assertThatCode(() -> {
            ReflectionUtils.setFieldValue(dto, intField, Integer.valueOf(42));
            ReflectionUtils.setFieldValue(dto, IntegerField, 84);
        }).doesNotThrowAnyException();

        assertThat(dto.primitiveInt).isEqualTo(42);
        assertThat(dto.wrapperInteger).isEqualTo(84);
    }

    public static class TestDtoWithSetter {
        private String name;
        private Integer age;
        public boolean setterCalled = false;

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
            this.setterCalled = true;
        }

        public Integer getAge() {
            return age;
        }

        public void setAge(Integer age) {
            this.age = age;
        }
    }

    public static class TestDtoWithGetter {
        private String name;
        public boolean getterCalled = false;

        public String getName() {
            this.getterCalled = true;
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }
    }

    public static class TestDtoWithBooleanGetter {
        private boolean active;
        public boolean isMethodCalled = false;

        public boolean isActive() {
            this.isMethodCalled = true;
            return active;
        }

        public void setActive(boolean active) {
            this.active = active;
        }
    }

    public static class TestDtoWithoutSetter {
        public String name;
    }

    public static class TestDtoWithoutGetter {
        public String name;
    }

    public static class TestDtoWithFinalField {
        private final String finalValue = "FINAL";
    }

    public static class TestDtoWithPrivateField {
        private String secretValue;
    }

    public static class TestDtoWithPrimitives {
        public int primitiveInt;
        public Integer wrapperInteger;
    }

    private Field getField(
            Class<?> clazz,
            String fieldName
    ) {
        try {
            return clazz.getDeclaredField(fieldName);
        } catch (NoSuchFieldException e) {
            throw new RuntimeException("Field not found: " + fieldName, e);
        }
    }
}