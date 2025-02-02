// Template Source: BaseEntityRequest.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests;

import com.microsoft.graph.http.IRequestBuilder;
import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.models.UserFlowLanguagePage;
import java.util.Arrays;
import java.util.EnumSet;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;
import com.microsoft.graph.core.IBaseClient;
import com.microsoft.graph.http.BaseRequest;
import com.microsoft.graph.http.HttpMethod;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the User Flow Language Page Request.
 */
public class UserFlowLanguagePageRequest extends BaseRequest<UserFlowLanguagePage> {
	
    /**
     * The request for the UserFlowLanguagePage
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     */
    public UserFlowLanguagePageRequest(@Nonnull final String requestUrl, @Nonnull final IBaseClient<?> client, @Nullable final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions) {
        super(requestUrl, client, requestOptions, UserFlowLanguagePage.class);
    }

    /**
     * Gets the UserFlowLanguagePage from the service
     *
     * @return a future with the result
     */
    @Nonnull
    public java.util.concurrent.CompletableFuture<UserFlowLanguagePage> getAsync() {
        return sendAsync(HttpMethod.GET, null);
    }

    /**
     * Gets the UserFlowLanguagePage from the service
     *
     * @return the UserFlowLanguagePage from the request
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public UserFlowLanguagePage get() throws ClientException {
       return send(HttpMethod.GET, null);
    }

    /**
     * Delete this item from the service
     *
     * @return a future with the deletion result
     */
    @Nonnull
    public java.util.concurrent.CompletableFuture<UserFlowLanguagePage> deleteAsync() {
        return sendAsync(HttpMethod.DELETE, null);
    }

    /**
     * Delete this item from the service
     * @return the resulting response if the service returns anything on deletion
     *
     * @throws ClientException if there was an exception during the delete operation
     */
    @Nullable
    public UserFlowLanguagePage delete() throws ClientException {
        return send(HttpMethod.DELETE, null);
    }

    /**
     * Patches this UserFlowLanguagePage with a source
     *
     * @param sourceUserFlowLanguagePage the source object with updates
     * @return a future with the result
     */
    @Nonnull
    public java.util.concurrent.CompletableFuture<UserFlowLanguagePage> patchAsync(@Nonnull final UserFlowLanguagePage sourceUserFlowLanguagePage) {
        return sendAsync(HttpMethod.PATCH, sourceUserFlowLanguagePage);
    }

    /**
     * Patches this UserFlowLanguagePage with a source
     *
     * @param sourceUserFlowLanguagePage the source object with updates
     * @return the updated UserFlowLanguagePage
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public UserFlowLanguagePage patch(@Nonnull final UserFlowLanguagePage sourceUserFlowLanguagePage) throws ClientException {
        return send(HttpMethod.PATCH, sourceUserFlowLanguagePage);
    }

    /**
     * Creates a UserFlowLanguagePage with a new object
     *
     * @param newUserFlowLanguagePage the new object to create
     * @return a future with the result
     */
    @Nonnull
    public java.util.concurrent.CompletableFuture<UserFlowLanguagePage> postAsync(@Nonnull final UserFlowLanguagePage newUserFlowLanguagePage) {
        return sendAsync(HttpMethod.POST, newUserFlowLanguagePage);
    }

    /**
     * Creates a UserFlowLanguagePage with a new object
     *
     * @param newUserFlowLanguagePage the new object to create
     * @return the created UserFlowLanguagePage
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public UserFlowLanguagePage post(@Nonnull final UserFlowLanguagePage newUserFlowLanguagePage) throws ClientException {
        return send(HttpMethod.POST, newUserFlowLanguagePage);
    }

    /**
     * Creates a UserFlowLanguagePage with a new object
     *
     * @param newUserFlowLanguagePage the object to create/update
     * @return a future with the result
     */
    @Nonnull
    public java.util.concurrent.CompletableFuture<UserFlowLanguagePage> putAsync(@Nonnull final UserFlowLanguagePage newUserFlowLanguagePage) {
        return sendAsync(HttpMethod.PUT, newUserFlowLanguagePage);
    }

    /**
     * Creates a UserFlowLanguagePage with a new object
     *
     * @param newUserFlowLanguagePage the object to create/update
     * @return the created UserFlowLanguagePage
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public UserFlowLanguagePage put(@Nonnull final UserFlowLanguagePage newUserFlowLanguagePage) throws ClientException {
        return send(HttpMethod.PUT, newUserFlowLanguagePage);
    }

    /**
     * Sets the select clause for the request
     *
     * @param value the select clause
     * @return the updated request
     */
     @Nonnull
     public UserFlowLanguagePageRequest select(@Nonnull final String value) {
         addSelectOption(value);
         return this;
     }

    /**
     * Sets the expand clause for the request
     *
     * @param value the expand clause
     * @return the updated request
     */
     @Nonnull
     public UserFlowLanguagePageRequest expand(@Nonnull final String value) {
         addExpandOption(value);
         return this;
     }

}

