package cn.qxl.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 忽略导出的字段
 *<li> @author: qiu</li>
 *<li>CreateDate:2018年10月10日</li>
 *<li>CreateTime:下午5:25:09</li>
 *<li>#IngnoreField</li>
 */
@Target({ ElementType.FIELD, ElementType.METHOD, ElementType.TYPE })
@Retention(RetentionPolicy.RUNTIME)
public @interface IngnoreField {
	
}
