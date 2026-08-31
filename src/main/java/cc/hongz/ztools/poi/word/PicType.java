package cc.hongz.ztools.poi.word;

import org.apache.poi.xwpf.usermodel.Document;

/**
 * Word中的图片类型。
 *
 * @author looly
 * @since 5.1.6
 */
public enum PicType {
    /**
     * Windows增强型图元文件（EMF）图片类型。
     */
    EMF(Document.PICTURE_TYPE_EMF),
    /**
     * Windows图元文件（WMF）图片类型。
     */
    WMF(Document.PICTURE_TYPE_WMF),
    /**
     * Macintosh PICT图片类型。
     */
    PICT(Document.PICTURE_TYPE_PICT),
    /**
     * JPEG图片类型。
     */
    JPEG(Document.PICTURE_TYPE_JPEG),
    /**
     * PNG图片类型。
     */
    PNG(Document.PICTURE_TYPE_PNG),
    /**
     * 设备无关位图（DIB）图片类型。
     */
    DIB(Document.PICTURE_TYPE_DIB),
    /**
     * GIF图片类型。
     */
    GIF(Document.PICTURE_TYPE_GIF),
    /**
     * TIFF图片类型。
     */
    TIFF(Document.PICTURE_TYPE_TIFF),
    /**
     * EPS图片类型。
     */
    EPS(Document.PICTURE_TYPE_EPS),
    /**
     * WPG图片类型。
     */
    WPG(Document.PICTURE_TYPE_WPG);

    /**
     * 构造。
     *
     * @param value 图片类型值
     */
    PicType(final int value) {
        this.value = value;
    }

    private final int value;

    /**
     * 获取图片类型对应值。
     *
     * @return 图片值
     */
    public int getValue() {
        return this.value;
    }
}
