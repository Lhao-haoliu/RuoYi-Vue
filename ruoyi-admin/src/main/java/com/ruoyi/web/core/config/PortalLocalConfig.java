package com.ruoyi.web.core.config;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

/**
 * Local portal configuration loaded from config/local-config.yml.
 */
@Component
public class PortalLocalConfig
{
    @Value("${commander.base-url:http://localhost:9000}")
    private String commanderBaseUrl;

    @Value("${portal.title:新系统门户}")
    private String portalTitle;

    public String getCommanderBaseUrl()
    {
        return commanderBaseUrl;
    }

    public String getPortalTitle()
    {
        return portalTitle;
    }
}
