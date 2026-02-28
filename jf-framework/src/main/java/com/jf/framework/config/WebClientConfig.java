package com.jf.framework.config;

import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.web.reactive.function.client.ExchangeStrategies;
import org.springframework.web.reactive.function.client.WebClient;

@Configuration
public class WebClientConfig {

    @Bean
    public WebClient webClient() {
        // 设置缓冲区大小，例如 16MB
        final int bufferSize = 16 * 1024 * 1024;  // 调整为合适大小，避免内存溢出
        ExchangeStrategies strategies = ExchangeStrategies.builder()
                .codecs(configurer -> configurer.defaultCodecs().maxInMemorySize(bufferSize))
                .build();

        return WebClient.builder()
                .exchangeStrategies(strategies)
                .baseUrl("http://localhost:9100")  // 根据您的服务调整
                .build();
    }
}