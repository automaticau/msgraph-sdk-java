// Template Source: BaseEntityRequest.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests.extensions;

import com.microsoft.graph.http.IRequestBuilder;
import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.concurrency.ICallback;
import com.microsoft.graph.models.extensions.MicrosoftStoreForBusinessApp;
import java.util.Arrays;
import java.util.EnumSet;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;
import com.microsoft.graph.core.IBaseClient;
import com.microsoft.graph.http.BaseRequest;
import com.microsoft.graph.http.HttpMethod;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Microsoft Store For Business App Request.
 */
public class MicrosoftStoreForBusinessAppRequest extends BaseRequest<MicrosoftStoreForBusinessApp> {
	
    /**
     * The request for the MicrosoftStoreForBusinessApp
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     */
    public MicrosoftStoreForBusinessAppRequest(@Nonnull final String requestUrl, @Nonnull final IBaseClient client, @Nullable final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions) {
        super(requestUrl, client, requestOptions, MicrosoftStoreForBusinessApp.class);
    }

    /**
     * Gets the MicrosoftStoreForBusinessApp from the service
     *
     * @param callback the callback to be called after success or failure
     */
    public void get(@Nonnull final ICallback<? super MicrosoftStoreForBusinessApp> callback) {
        send(HttpMethod.GET, callback, null);
    }

    /**
     * Gets the MicrosoftStoreForBusinessApp from the service
     *
     * @return the MicrosoftStoreForBusinessApp from the request
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public MicrosoftStoreForBusinessApp get() throws ClientException {
       return send(HttpMethod.GET, null);
    }

    /**
     * Delete this item from the service
     *
     * @param callback the callback when the deletion action has completed
     */
    public void delete(@Nonnull final ICallback<? super MicrosoftStoreForBusinessApp> callback) {
        send(HttpMethod.DELETE, callback, null);
    }

    /**
     * Delete this item from the service
     *
     * @throws ClientException if there was an exception during the delete operation
     */
    public void delete() throws ClientException {
        send(HttpMethod.DELETE, null);
    }

    /**
     * Patches this MicrosoftStoreForBusinessApp with a source
     *
     * @param sourceMicrosoftStoreForBusinessApp the source object with updates
     * @param callback the callback to be called after success or failure
     */
    public void patch(@Nonnull final MicrosoftStoreForBusinessApp sourceMicrosoftStoreForBusinessApp, @Nonnull final ICallback<? super MicrosoftStoreForBusinessApp> callback) {
        send(HttpMethod.PATCH, callback, sourceMicrosoftStoreForBusinessApp);
    }

    /**
     * Patches this MicrosoftStoreForBusinessApp with a source
     *
     * @param sourceMicrosoftStoreForBusinessApp the source object with updates
     * @return the updated MicrosoftStoreForBusinessApp
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public MicrosoftStoreForBusinessApp patch(@Nonnull final MicrosoftStoreForBusinessApp sourceMicrosoftStoreForBusinessApp) throws ClientException {
        return send(HttpMethod.PATCH, sourceMicrosoftStoreForBusinessApp);
    }

    /**
     * Creates a MicrosoftStoreForBusinessApp with a new object
     *
     * @param newMicrosoftStoreForBusinessApp the new object to create
     * @param callback the callback to be called after success or failure
     */
    public void post(@Nonnull final MicrosoftStoreForBusinessApp newMicrosoftStoreForBusinessApp, @Nonnull final ICallback<? super MicrosoftStoreForBusinessApp> callback) {
        send(HttpMethod.POST, callback, newMicrosoftStoreForBusinessApp);
    }

    /**
     * Creates a MicrosoftStoreForBusinessApp with a new object
     *
     * @param newMicrosoftStoreForBusinessApp the new object to create
     * @return the created MicrosoftStoreForBusinessApp
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public MicrosoftStoreForBusinessApp post(@Nonnull final MicrosoftStoreForBusinessApp newMicrosoftStoreForBusinessApp) throws ClientException {
        return send(HttpMethod.POST, newMicrosoftStoreForBusinessApp);
    }

    /**
     * Creates a MicrosoftStoreForBusinessApp with a new object
     *
     * @param newMicrosoftStoreForBusinessApp the object to create/update
     * @param callback the callback to be called after success or failure
     */
    public void put(@Nonnull final MicrosoftStoreForBusinessApp newMicrosoftStoreForBusinessApp, @Nonnull final ICallback<? super MicrosoftStoreForBusinessApp> callback) {
        send(HttpMethod.PUT, callback, newMicrosoftStoreForBusinessApp);
    }

    /**
     * Creates a MicrosoftStoreForBusinessApp with a new object
     *
     * @param newMicrosoftStoreForBusinessApp the object to create/update
     * @return the created MicrosoftStoreForBusinessApp
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public MicrosoftStoreForBusinessApp put(@Nonnull final MicrosoftStoreForBusinessApp newMicrosoftStoreForBusinessApp) throws ClientException {
        return send(HttpMethod.PUT, newMicrosoftStoreForBusinessApp);
    }

    /**
     * Sets the select clause for the request
     *
     * @param value the select clause
     * @return the updated request
     */
     @Nonnull
     public MicrosoftStoreForBusinessAppRequest select(@Nonnull final String value) {
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
     public MicrosoftStoreForBusinessAppRequest expand(@Nonnull final String value) {
         addExpandOption(value);
         return this;
     }

}

