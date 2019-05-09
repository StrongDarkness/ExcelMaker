package cn.qxl.annotation;

import java.lang.annotation.*;

/**
 * 修改字段注解<br>
 * 注解到bean上<br>
 * 该方法会自动修改匹配的字段<br>
 * 该方法只用于String类型字段
 * @author qiu
 *
 */
@Target({ ElementType.FIELD, ElementType.METHOD, ElementType.TYPE })
@Retention(RetentionPolicy.RUNTIME)
@Repeatable(StateFields.class)
public @interface StateField {
	/**
	 * 需要匹配的字段
	 * @return
	 */
	String key();
	/**
	 * 需要修改的值(默认为空字符)
	 * @return
	 */
	String value() default "";
	
}
