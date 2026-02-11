package com.jf.common.exception.file;

import com.jf.common.exception.base.BaseException;

/**
 * 文件信息异常类
 * 
 * @author jf
 */
public class FileException extends BaseException
{
    private static final long serialVersionUID = 1L;

    public FileException(String code, Object[] args)
    {
        super("file", code, args, null);
    }

}
