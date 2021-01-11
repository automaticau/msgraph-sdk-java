// Template Source: BaseEntityRequest.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests.extensions;

import com.microsoft.graph.http.IRequestBuilder;
import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.concurrency.ICallback;
import com.microsoft.graph.models.extensions.IosVppApp;
import java.util.Arrays;
import java.util.EnumSet;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;
import com.microsoft.graph.core.IBaseClient;
import com.microsoft.graph.http.BaseRequest;
import com.microsoft.graph.http.HttpMethod;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Ios Vpp App Request.
 */
public class IosVppAppRequest extends BaseRequest<IosVppApp> {
	
    /**
     * The request for the IosVppApp
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     */
    public IosVppAppRequest(@Nonnull final String requestUrl, @Nonnull final IBaseClient client, @Nullable final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions) {
        super(requestUrl, client, requestOptions, IosVppApp.class);
    }

    /**
     * Gets the IosVppApp from the service
     *
     * @param callback the callback to be called after success or failure
     */
    public void get(@Nonnull final ICallback<? super IosVppApp> callback) {
        send(HttpMethod.GET, callback, null);
    }

    /**
     * Gets the IosVppApp from the service
     *
     * @return the IosVppApp from the request
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public IosVppApp get() throws ClientException {
       return send(HttpMethod.GET, null);
    }

    /**
     * Delete this item from the service
     *
     * @param callback the callback when the deletion action has completed
     */
    public void delete(@Nonnull final ICallback<? super IosVppApp> callback) {
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
     * Patches this IosVppApp with a source
     *
     * @param sourceIosVppApp the source object with updates
     * @param callback the callback to be called after success or failure
     */
    public void patch(@Nonnull final IosVppApp sourceIosVppApp, @Nonnull final ICallback<? super IosVppApp> callback) {
        send(HttpMethod.PATCH, callback, sourceIosVppApp);
    }

    /**
     * Patches this IosVppApp with a source
     *
     * @param sourceIosVppApp the source object with updates
     * @return the updated IosVppApp
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public IosVppApp patch(@Nonnull final IosVppApp sourceIosVppApp) throws ClientException {
        return send(HttpMethod.PATCH, sourceIosVppApp);
    }

    /**
     * Creates a IosVppApp with a new object
     *
     * @param newIosVppApp the new object to create
     * @param callback the callback to be called after success or failure
     */
    public void post(@Nonnull final IosVppApp newIosVppApp, @Nonnull final ICallback<? super IosVppApp> callback) {
        send(HttpMethod.POST, callback, newIosVppApp);
    }

    /**
     * Creates a IosVppApp with a new object
     *
     * @param newIosVppApp the new object to create
     * @return the created IosVppApp
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public IosVppApp post(@Nonnull final IosVppApp newIosVppApp) throws ClientException {
        return send(HttpMethod.POST, newIosVppApp);
    }

    /**
     * Creates a IosVppApp with a new object
     *
     * @param newIosVppApp the object to create/update
     * @param callback the callback to be called after success or failure
     */
    public void put(@Nonnull final IosVppApp newIosVppApp, @Nonnull final ICallback<? super IosVppApp> callback) {
        send(HttpMethod.PUT, callback, newIosVppApp);
    }

    /**
     * Creates a IosVppApp with a new object
     *
     * @param newIosVppApp the object to create/update
     * @return the created IosVppApp
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public IosVppApp put(@Nonnull final IosVppApp newIosVppApp) throws ClientException {
        return send(HttpMethod.PUT, newIosVppApp);
    }

    /**
     * Sets the select clause for the request
     *
     * @param value the select clause
     * @return the updated request
     */
     @Nonnull
     public IosVppAppRequest select(@Nonnull final String value) {
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
     public IosVppAppRequest expand(@Nonnull final String value) {
         addExpandOption(value);
         return this;
     }

}

