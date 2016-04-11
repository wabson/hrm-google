package org.springframework.social.quickstart.user;

import org.springframework.social.connect.Connection;
import org.springframework.social.connect.ConnectionFactory;
import org.springframework.social.connect.web.ProviderSignInInterceptor;
import org.springframework.social.google.api.Google;
import org.springframework.util.MultiValueMap;
import org.springframework.web.context.request.WebRequest;

public class IncludeEmailProviderSignInInterceptor implements ProviderSignInInterceptor<Google> {

	@Override
	public void preSignIn(ConnectionFactory<Google> connectionFactory,
			MultiValueMap<String, String> parameters, WebRequest request) {
		String login = request.getParameter("login");
		if (login != null && !login.equals("")) {
			parameters.add("login_hint", login);
		}
	}

	@Override
	public void postSignIn(Connection<Google> connection, WebRequest request) {
		// Do nothing
	}

}
