// Template Source: BaseEntityRequest.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests;

import com.microsoft.graph.http.IRequestBuilder;
import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.models.WindowsHelloForBusinessAuthenticationMethod;
import com.microsoft.graph.requests.DeviceRequestBuilder;
import java.util.Arrays;
import java.util.EnumSet;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;
import com.microsoft.graph.core.IBaseClient;
import com.microsoft.graph.http.BaseRequest;
import com.microsoft.graph.http.HttpMethod;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Windows Hello For Business Authentication Method Request.
 */
public class WindowsHelloForBusinessAuthenticationMethodRequest extends BaseRequest<WindowsHelloForBusinessAuthenticationMethod> {
	
    /**
     * The request for the WindowsHelloForBusinessAuthenticationMethod
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     */
    public WindowsHelloForBusinessAuthenticationMethodRequest(@Nonnull final String requestUrl, @Nonnull final IBaseClient<?> client, @Nullable final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions) {
        super(requestUrl, client, requestOptions, WindowsHelloForBusinessAuthenticationMethod.class);
    }

    /**
     * Gets the WindowsHelloForBusinessAuthenticationMethod from the service
     *
     * @return a future with the result
     */
    @Nonnull
    public java.util.concurrent.CompletableFuture<WindowsHelloForBusinessAuthenticationMethod> getAsync() {
        return sendAsync(HttpMethod.GET, null);
    }

    /**
     * Gets the WindowsHelloForBusinessAuthenticationMethod from the service
     *
     * @return the WindowsHelloForBusinessAuthenticationMethod from the request
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public WindowsHelloForBusinessAuthenticationMethod get() throws ClientException {
       return send(HttpMethod.GET, null);
    }

    /**
     * Delete this item from the service
     *
     * @return a future with the deletion result
     */
    @Nonnull
    public java.util.concurrent.CompletableFuture<WindowsHelloForBusinessAuthenticationMethod> deleteAsync() {
        return sendAsync(HttpMethod.DELETE, null);
    }

    /**
     * Delete this item from the service
     * @return the resulting response if the service returns anything on deletion
     *
     * @throws ClientException if there was an exception during the delete operation
     */
    @Nullable
    public WindowsHelloForBusinessAuthenticationMethod delete() throws ClientException {
        return send(HttpMethod.DELETE, null);
    }

    /**
     * Patches this WindowsHelloForBusinessAuthenticationMethod with a source
     *
     * @param sourceWindowsHelloForBusinessAuthenticationMethod the source object with updates
     * @return a future with the result
     */
    @Nonnull
    public java.util.concurrent.CompletableFuture<WindowsHelloForBusinessAuthenticationMethod> patchAsync(@Nonnull final WindowsHelloForBusinessAuthenticationMethod sourceWindowsHelloForBusinessAuthenticationMethod) {
        return sendAsync(HttpMethod.PATCH, sourceWindowsHelloForBusinessAuthenticationMethod);
    }

    /**
     * Patches this WindowsHelloForBusinessAuthenticationMethod with a source
     *
     * @param sourceWindowsHelloForBusinessAuthenticationMethod the source object with updates
     * @return the updated WindowsHelloForBusinessAuthenticationMethod
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public WindowsHelloForBusinessAuthenticationMethod patch(@Nonnull final WindowsHelloForBusinessAuthenticationMethod sourceWindowsHelloForBusinessAuthenticationMethod) throws ClientException {
        return send(HttpMethod.PATCH, sourceWindowsHelloForBusinessAuthenticationMethod);
    }

    /**
     * Creates a WindowsHelloForBusinessAuthenticationMethod with a new object
     *
     * @param newWindowsHelloForBusinessAuthenticationMethod the new object to create
     * @return a future with the result
     */
    @Nonnull
    public java.util.concurrent.CompletableFuture<WindowsHelloForBusinessAuthenticationMethod> postAsync(@Nonnull final WindowsHelloForBusinessAuthenticationMethod newWindowsHelloForBusinessAuthenticationMethod) {
        return sendAsync(HttpMethod.POST, newWindowsHelloForBusinessAuthenticationMethod);
    }

    /**
     * Creates a WindowsHelloForBusinessAuthenticationMethod with a new object
     *
     * @param newWindowsHelloForBusinessAuthenticationMethod the new object to create
     * @return the created WindowsHelloForBusinessAuthenticationMethod
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public WindowsHelloForBusinessAuthenticationMethod post(@Nonnull final WindowsHelloForBusinessAuthenticationMethod newWindowsHelloForBusinessAuthenticationMethod) throws ClientException {
        return send(HttpMethod.POST, newWindowsHelloForBusinessAuthenticationMethod);
    }

    /**
     * Creates a WindowsHelloForBusinessAuthenticationMethod with a new object
     *
     * @param newWindowsHelloForBusinessAuthenticationMethod the object to create/update
     * @return a future with the result
     */
    @Nonnull
    public java.util.concurrent.CompletableFuture<WindowsHelloForBusinessAuthenticationMethod> putAsync(@Nonnull final WindowsHelloForBusinessAuthenticationMethod newWindowsHelloForBusinessAuthenticationMethod) {
        return sendAsync(HttpMethod.PUT, newWindowsHelloForBusinessAuthenticationMethod);
    }

    /**
     * Creates a WindowsHelloForBusinessAuthenticationMethod with a new object
     *
     * @param newWindowsHelloForBusinessAuthenticationMethod the object to create/update
     * @return the created WindowsHelloForBusinessAuthenticationMethod
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public WindowsHelloForBusinessAuthenticationMethod put(@Nonnull final WindowsHelloForBusinessAuthenticationMethod newWindowsHelloForBusinessAuthenticationMethod) throws ClientException {
        return send(HttpMethod.PUT, newWindowsHelloForBusinessAuthenticationMethod);
    }

    /**
     * Sets the select clause for the request
     *
     * @param value the select clause
     * @return the updated request
     */
     @Nonnull
     public WindowsHelloForBusinessAuthenticationMethodRequest select(@Nonnull final String value) {
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
     public WindowsHelloForBusinessAuthenticationMethodRequest expand(@Nonnull final String value) {
         addExpandOption(value);
         return this;
     }

}
