package cc.hongz.ztools.poi.exceptions;

import cc.hongz.ztools.core.exceptions.ExceptionUtil;
import cc.hongz.ztools.core.util.StrUtil;

/**
 * POI异常。
 *
 * @author xiaoleilu
 */
public class POIException extends RuntimeException {
    private static final long serialVersionUID = 2711633732613506552L;

    /**
     * 构造。
     *
     * @param e 异常
     */
    public POIException(final Throwable e) {
        super(ExceptionUtil.getMessage(e), e);
    }

    /**
     * 构造。
     *
     * @param message 错误消息
     */
    public POIException(final String message) {
        super(message);
    }

    /**
     * 构造。
     *
     * @param messageTemplate 消息模板
     * @param params          参数
     */
    public POIException(final String messageTemplate, final Object... params) {
        super(StrUtil.format(messageTemplate, params));
    }

    /**
     * 构造。
     *
     * @param message   错误消息
     * @param throwable 异常
     */
    public POIException(final String message, final Throwable throwable) {
        super(message, throwable);
    }

    /**
     * 构造。
     *
     * @param message            错误消息
     * @param throwable          异常
     * @param enableSuppression  是否启用异常抑制
     * @param writableStackTrace 是否启用栈轨迹
     */
    public POIException(final String message, final Throwable throwable, final boolean enableSuppression,
            final boolean writableStackTrace) {
        super(message, throwable, enableSuppression, writableStackTrace);
    }

    /**
     * 构造。
     *
     * @param throwable       异常
     * @param messageTemplate 消息模板
     * @param params          参数
     */
    public POIException(final Throwable throwable, final String messageTemplate, final Object... params) {
        super(StrUtil.format(messageTemplate, params), throwable);
    }
}
