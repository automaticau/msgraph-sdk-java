// Template Source: BaseEntity.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.models;
import com.microsoft.graph.serializer.ISerializer;
import com.microsoft.graph.serializer.IJsonBackedObject;
import com.microsoft.graph.serializer.AdditionalDataManager;
import java.util.EnumSet;
import com.microsoft.graph.models.StateManagementSetting;


import com.google.gson.JsonObject;
import com.google.gson.annotations.SerializedName;
import com.google.gson.annotations.Expose;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Windows Firewall Network Profile.
 */
public class WindowsFirewallNetworkProfile implements IJsonBackedObject {

    /** the OData type of the object as returned by the service */
    @SerializedName("@odata.type")
    @Expose
    @Nullable
    public String oDataType;

    private transient AdditionalDataManager additionalDataManager = new AdditionalDataManager(this);

    @Override
    @Nonnull
    public final AdditionalDataManager additionalDataManager() {
        return additionalDataManager;
    }

    /**
     * The Authorized Application Rules From Group Policy Merged.
     * Configures the firewall to merge authorized application rules from group policy with those from local store instead of ignoring the local store rules. When AuthorizedApplicationRulesFromGroupPolicyNotMerged and AuthorizedApplicationRulesFromGroupPolicyMerged are both true, AuthorizedApplicationRulesFromGroupPolicyMerged takes priority.
     */
    @SerializedName(value = "authorizedApplicationRulesFromGroupPolicyMerged", alternate = {"AuthorizedApplicationRulesFromGroupPolicyMerged"})
    @Expose
	@Nullable
    public Boolean authorizedApplicationRulesFromGroupPolicyMerged;

    /**
     * The Connection Security Rules From Group Policy Merged.
     * Configures the firewall to merge connection security rules from group policy with those from local store instead of ignoring the local store rules. When ConnectionSecurityRulesFromGroupPolicyNotMerged and ConnectionSecurityRulesFromGroupPolicyMerged are both true, ConnectionSecurityRulesFromGroupPolicyMerged takes priority.
     */
    @SerializedName(value = "connectionSecurityRulesFromGroupPolicyMerged", alternate = {"ConnectionSecurityRulesFromGroupPolicyMerged"})
    @Expose
	@Nullable
    public Boolean connectionSecurityRulesFromGroupPolicyMerged;

    /**
     * The Firewall Enabled.
     * Configures the host device to allow or block the firewall and advanced security enforcement for the network profile. Possible values are: notConfigured, blocked, allowed.
     */
    @SerializedName(value = "firewallEnabled", alternate = {"FirewallEnabled"})
    @Expose
	@Nullable
    public StateManagementSetting firewallEnabled;

    /**
     * The Global Port Rules From Group Policy Merged.
     * Configures the firewall to merge global port rules from group policy with those from local store instead of ignoring the local store rules. When GlobalPortRulesFromGroupPolicyNotMerged and GlobalPortRulesFromGroupPolicyMerged are both true, GlobalPortRulesFromGroupPolicyMerged takes priority.
     */
    @SerializedName(value = "globalPortRulesFromGroupPolicyMerged", alternate = {"GlobalPortRulesFromGroupPolicyMerged"})
    @Expose
	@Nullable
    public Boolean globalPortRulesFromGroupPolicyMerged;

    /**
     * The Inbound Connections Blocked.
     * Configures the firewall to block all incoming connections by default. When InboundConnectionsRequired and InboundConnectionsBlocked are both true, InboundConnectionsBlocked takes priority.
     */
    @SerializedName(value = "inboundConnectionsBlocked", alternate = {"InboundConnectionsBlocked"})
    @Expose
	@Nullable
    public Boolean inboundConnectionsBlocked;

    /**
     * The Inbound Notifications Blocked.
     * Prevents the firewall from displaying notifications when an application is blocked from listening on a port. When InboundNotificationsRequired and InboundNotificationsBlocked are both true, InboundNotificationsBlocked takes priority.
     */
    @SerializedName(value = "inboundNotificationsBlocked", alternate = {"InboundNotificationsBlocked"})
    @Expose
	@Nullable
    public Boolean inboundNotificationsBlocked;

    /**
     * The Incoming Traffic Blocked.
     * Configures the firewall to block all incoming traffic regardless of other policy settings. When IncomingTrafficRequired and IncomingTrafficBlocked are both true, IncomingTrafficBlocked takes priority.
     */
    @SerializedName(value = "incomingTrafficBlocked", alternate = {"IncomingTrafficBlocked"})
    @Expose
	@Nullable
    public Boolean incomingTrafficBlocked;

    /**
     * The Outbound Connections Blocked.
     * Configures the firewall to block all outgoing connections by default. When OutboundConnectionsRequired and OutboundConnectionsBlocked are both true, OutboundConnectionsBlocked takes priority. This setting will get applied to Windows releases version 1809 and above.
     */
    @SerializedName(value = "outboundConnectionsBlocked", alternate = {"OutboundConnectionsBlocked"})
    @Expose
	@Nullable
    public Boolean outboundConnectionsBlocked;

    /**
     * The Policy Rules From Group Policy Merged.
     * Configures the firewall to merge Firewall Rule policies from group policy with those from local store instead of ignoring the local store rules. When PolicyRulesFromGroupPolicyNotMerged and PolicyRulesFromGroupPolicyMerged are both true, PolicyRulesFromGroupPolicyMerged takes priority.
     */
    @SerializedName(value = "policyRulesFromGroupPolicyMerged", alternate = {"PolicyRulesFromGroupPolicyMerged"})
    @Expose
	@Nullable
    public Boolean policyRulesFromGroupPolicyMerged;

    /**
     * The Secured Packet Exemption Allowed.
     * Configures the firewall to allow the host computer to respond to unsolicited network traffic of that traffic is secured by IPSec even when stealthModeBlocked is set to true. When SecuredPacketExemptionBlocked and SecuredPacketExemptionAllowed are both true, SecuredPacketExemptionAllowed takes priority.
     */
    @SerializedName(value = "securedPacketExemptionAllowed", alternate = {"SecuredPacketExemptionAllowed"})
    @Expose
	@Nullable
    public Boolean securedPacketExemptionAllowed;

    /**
     * The Stealth Mode Blocked.
     * Prevent the server from operating in stealth mode. When StealthModeRequired and StealthModeBlocked are both true, StealthModeBlocked takes priority.
     */
    @SerializedName(value = "stealthModeBlocked", alternate = {"StealthModeBlocked"})
    @Expose
	@Nullable
    public Boolean stealthModeBlocked;

    /**
     * The Unicast Responses To Multicast Broadcasts Blocked.
     * Configures the firewall to block unicast responses to multicast broadcast traffic. When UnicastResponsesToMulticastBroadcastsRequired and UnicastResponsesToMulticastBroadcastsBlocked are both true, UnicastResponsesToMulticastBroadcastsBlocked takes priority.
     */
    @SerializedName(value = "unicastResponsesToMulticastBroadcastsBlocked", alternate = {"UnicastResponsesToMulticastBroadcastsBlocked"})
    @Expose
	@Nullable
    public Boolean unicastResponsesToMulticastBroadcastsBlocked;


    /**
     * Sets the raw JSON object
     *
     * @param serializer the serializer
     * @param json the JSON object to set this object to
     */
    public void setRawObject(@Nonnull final ISerializer serializer, @Nonnull final JsonObject json) {

    }
}
