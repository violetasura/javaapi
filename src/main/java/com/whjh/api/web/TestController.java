package com.whjh.api.web;

import java.util.Map;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

@Controller
@RequestMapping(value = "/hello")
public class TestController {

	@RequestMapping("/helloworld")
	public String helloworld() {

		return "helloworld";
	}

	@GetMapping("/helloworld2")
	public String helloworld2() {

		return "abcefg";
	}

	@RequestMapping("/index")
	public ModelAndView index(Map<String, Object> map) {
		System.out.println("HelloController.helloJsp().hello=");

		map.put("hello", "aaa");
		ModelAndView modelAndView = new ModelAndView("mamabi");

		return modelAndView;

	}

}
