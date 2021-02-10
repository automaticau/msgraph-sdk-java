// Template Source: IBaseEntityRequest.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests.extensions;
import com.microsoft.graph.models.extensions.Fido2AuthenticationMethodConfiguration;

import com.microsoft.graph.concurrency.ICallback;
import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.http.IHttpRequest;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The interface for the Fido2Authentication Method Configuration Request.
 */
public interface IFido2AuthenticationMethodConfigurationRequest extends IHttpRequest {

    /**
     * Gets the Fido2AuthenticationMethodConfiguration from the service
     *
     * @param callback the callback to be called after success or failure
     */
    void get(final ICallback<? super Fido2AuthenticationMethodConfiguration> callback);

    /**
     * Gets the Fido2AuthenticationMethodConfiguration from the service
     *
     * @return the Fido2AuthenticationMethodConfiguration from the request
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    Fido2AuthenticationMethodConfiguration get() throws ClientException;

    /**
     * Delete this item from the service
     *
     * @param callback the callback when the deletion action has completed
     */
    void delete(final ICallback<? super Fido2AuthenticationMethodConfiguration> callback);

    /**
     * Delete this item from the service
     *
     * @throws ClientException if there was an exception during the delete operation
     */
    void delete() throws ClientException;

    /**
     * Patches this Fido2AuthenticationMethodConfiguration with a source
     *
     * @param sourceFido2AuthenticationMethodConfiguration the source object with updates
     * @param callback the callback to be called after success or failure
     */
    void patch(final Fido2AuthenticationMethodConfiguration sourceFido2AuthenticationMethodConfiguration, final ICallback<? super Fido2AuthenticationMethodConfiguration> callback);

    /**
     * Patches this Fido2AuthenticationMethodConfiguration with a source
     *
     * @param sourceFido2AuthenticationMethodConfiguration the source object with updates
     * @return the updated Fido2AuthenticationMethodConfiguration
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    Fido2AuthenticationMethodConfiguration patch(final Fido2AuthenticationMethodConfiguration sourceFido2AuthenticationMethodConfiguration) throws ClientException;

    /**
     * Posts a Fido2AuthenticationMethodConfiguration with a new object
     *
     * @param newFido2AuthenticationMethodConfiguration the new object to create
     * @param callback the callback to be called after success or failure
     */
    void post(final Fido2AuthenticationMethodConfiguration newFido2AuthenticationMethodConfiguration, final ICallback<? super Fido2AuthenticationMethodConfiguration> callback);

    /**
     * Posts a Fido2AuthenticationMethodConfiguration with a new object
     *
     * @param newFido2AuthenticationMethodConfiguration the new object to create
     * @return the created Fido2AuthenticationMethodConfiguration
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    Fido2AuthenticationMethodConfiguration post(final Fido2AuthenticationMethodConfiguration newFido2AuthenticationMethodConfiguration) throws ClientException;

    /**
     * Posts a Fido2AuthenticationMethodConfiguration with a new object
     *
     * @param newFido2AuthenticationMethodConfiguration the object to create/update
     * @param callback the callback to be called after success or failure
     */
    void put(final Fido2AuthenticationMethodConfiguration newFido2AuthenticationMethodConfiguration, final ICallback<? super Fido2AuthenticationMethodConfiguration> callback);

    /**
     * Posts a Fido2AuthenticationMethodConfiguration with a new object
     *
     * @param newFido2AuthenticationMethodConfiguration the object to create/update
     * @return the created Fido2AuthenticationMethodConfiguration
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    Fido2AuthenticationMethodConfiguration put(final Fido2AuthenticationMethodConfiguration newFido2AuthenticationMethodConfiguration) throws ClientException;

    /**
     * Sets the select clause for the request
     *
     * @param value the select clause
     * @return the updated request
     */
    IFido2AuthenticationMethodConfigurationRequest select(final String value);

    /**
     * Sets the expand clause for the request
     *
     * @param value the expand clause
     * @return the updated request
     */
    IFido2AuthenticationMethodConfigurationRequest expand(final String value);

}

