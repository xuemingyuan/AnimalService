package com.animal.utils;

public class StringUtil {
	public static String getFixedLenString(String str, int len, char c) {
		if (str == null || str.length() == 0){
			str = "";
		}
		if (str.length() == len){
			return str;
		}
		if (str.length() > len){
			return str.substring(0,len);
		}
		StringBuilder sb = new StringBuilder(str);
		while (sb.length() < len){
			sb.append(c);
		}
		return sb.toString();
	}
	
	public static void main(String[] args) {
		System.out.println(getFixedLenString("w",3,' '));
	}

	public static boolean isNullOrEmpty(String str) {
		return null == str || str.isEmpty();
	}
}
