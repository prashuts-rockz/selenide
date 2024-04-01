package org.tests;

import com.codeborne.selenide.WebDriverRunner;
import org.testng.annotations.Listeners;
import org.testng.annotations.Test;
import static com.codeborne.selenide.Selenide.*;
import static org.assertj.core.api.Assertions.*;
import com.codeborne.selenide.testng.ScreenShooter;

@Listeners({ ScreenShooter.class})
public class LoginPageTests {


    @Test
    public void verifyUrlAndTitle(){
        open("http://localhost/website/index.html");
        String url = WebDriverRunner.url();
        assertThat(url).as("Failed to assert URL ").isEqualTo("http://localhost/website/index.html");

        String title = WebDriverRunner.getWebDriver().getTitle();
        assertThat(title).as("Failed to assert title ").isEqualTo("MY SITE");
    }


}
