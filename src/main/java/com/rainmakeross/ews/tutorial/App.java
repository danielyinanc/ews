package com.rainmakeross.ews.tutorial;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.ApplicationContext;
import org.springframework.context.annotation.Bean;

@SpringBootApplication
public class App {
    @Value("${exchange.email}")
    private String exchangeEmail;

    @Value("${exchange.password}")
    private String exchangePassword;

    public static void main(String[] args) {
        SpringApplication.run(App.class, args);
    }

    @Bean
    public CommandLineRunner commandLineRunner(ApplicationContext ctx) {
        return args -> {
            pageThroughEntireInbox();
        };
    }

    public void pageThroughEntireInbox() throws Exception {
        int offset = 50;
        ExchangeService service = new ExchangeService();
        ExchangeCredentials credentials = new WebCredentials(exchangeEmail, exchangePassword);
        service.setCredentials(credentials);
        service.autodiscoverUrl(exchangeEmail);
        ItemView view = new ItemView(50);
        FindItemsResults<Item> findResults;
        findResults = service.findItems(WellKnownFolderName.Inbox, view);
        for (Item item : findResults.getItems()) {
            System.out.println(item.getSubject());
        }

    }
}
