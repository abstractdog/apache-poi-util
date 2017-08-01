package com.abstractdog.regex.test;

import java.util.regex.Pattern;

import org.junit.Test;

import org.junit.Assert;

public class RegexTest {

	@Test
	public void testChangePattern(){
		String text = "asdf 2345 ertt WRTRT";
		String from = "[a-z]";
		String to = "*";
		
		String result = Pattern.compile(from).matcher(text).replaceAll(to);
		Assert.assertEquals("**** 2345 **** WRTRT", result);
	}
}
