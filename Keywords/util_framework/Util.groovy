package util_framework

import com.kms.katalon.core.annotation.Keyword
import com.kms.katalon.core.model.FailureHandling
import com.kms.katalon.core.testobject.TestObject
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI


class Util {

	@Keyword
	static void highlightElement(TestObject to) {
		WebUI.executeJavaScript(
				'arguments[0].style.border="3px solid red"',
				Arrays.asList(WebUI.findWebElement(to, 10))
				)
	}

	@Keyword
	static void waitAndClick(TestObject to, int timeout = 10) {
		WebUI.waitForElementClickable(to, timeout)
		WebUI.click(to)
	}

	@Keyword
	static void setEncryptedTextIfVisible(TestObject to, String encryptedText) {
		if (WebUI.waitForElementVisible(to, 5, FailureHandling.OPTIONAL)) {
			WebUI.setEncryptedText(to, encryptedText)
		}
	}

	@Keyword
	static String getCurrentDate(String format = "yyyy-MM-dd") {
		return new Date().format(format)
	}

	@Keyword
	static void scrollToBottom() {
		WebUI.executeJavaScript("window.scrollTo(0, document.body.scrollHeight)", null)
	}
}
