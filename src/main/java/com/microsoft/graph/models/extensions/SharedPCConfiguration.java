// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.models.extensions;
import com.microsoft.graph.serializer.ISerializer;
import com.microsoft.graph.serializer.IJsonBackedObject;
import com.microsoft.graph.serializer.AdditionalDataManager;
import java.util.EnumSet;
import com.microsoft.graph.models.extensions.SharedPCAccountManagerPolicy;
import com.microsoft.graph.models.generated.SharedPCAllowedAccountType;
import com.microsoft.graph.models.extensions.DeviceConfiguration;


import com.google.gson.JsonObject;
import com.google.gson.annotations.SerializedName;
import com.google.gson.annotations.Expose;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Shared PCConfiguration.
 */
public class SharedPCConfiguration extends DeviceConfiguration implements IJsonBackedObject {


    /**
     * The Account Manager Policy.
     * Specifies how accounts are managed on a shared PC. Only applies when disableAccountManager is false.
     */
    @SerializedName(value = "accountManagerPolicy", alternate = {"AccountManagerPolicy"})
    @Expose
    public SharedPCAccountManagerPolicy accountManagerPolicy;

    /**
     * The Allowed Accounts.
     * Indicates which type of accounts are allowed to use on a shared PC. Possible values are: guest, domain.
     */
    @SerializedName(value = "allowedAccounts", alternate = {"AllowedAccounts"})
    @Expose
    public EnumSet<SharedPCAllowedAccountType> allowedAccounts;

    /**
     * The Allow Local Storage.
     * Specifies whether local storage is allowed on a shared PC.
     */
    @SerializedName(value = "allowLocalStorage", alternate = {"AllowLocalStorage"})
    @Expose
    public Boolean allowLocalStorage;

    /**
     * The Disable Account Manager.
     * Disables the account manager for shared PC mode.
     */
    @SerializedName(value = "disableAccountManager", alternate = {"DisableAccountManager"})
    @Expose
    public Boolean disableAccountManager;

    /**
     * The Disable Edu Policies.
     * Specifies whether the default shared PC education environment policies should be disabled. For Windows 10 RS2 and later, this policy will be applied without setting Enabled to true.
     */
    @SerializedName(value = "disableEduPolicies", alternate = {"DisableEduPolicies"})
    @Expose
    public Boolean disableEduPolicies;

    /**
     * The Disable Power Policies.
     * Specifies whether the default shared PC power policies should be disabled.
     */
    @SerializedName(value = "disablePowerPolicies", alternate = {"DisablePowerPolicies"})
    @Expose
    public Boolean disablePowerPolicies;

    /**
     * The Disable Sign In On Resume.
     * Disables the requirement to sign in whenever the device wakes up from sleep mode.
     */
    @SerializedName(value = "disableSignInOnResume", alternate = {"DisableSignInOnResume"})
    @Expose
    public Boolean disableSignInOnResume;

    /**
     * The Enabled.
     * Enables shared PC mode and applies the shared pc policies.
     */
    @SerializedName(value = "enabled", alternate = {"Enabled"})
    @Expose
    public Boolean enabled;

    /**
     * The Idle Time Before Sleep In Seconds.
     * Specifies the time in seconds that a device must sit idle before the PC goes to sleep. Setting this value to 0 prevents the sleep timeout from occurring.
     */
    @SerializedName(value = "idleTimeBeforeSleepInSeconds", alternate = {"IdleTimeBeforeSleepInSeconds"})
    @Expose
    public Integer idleTimeBeforeSleepInSeconds;

    /**
     * The Kiosk App Display Name.
     * Specifies the display text for the account shown on the sign-in screen which launches the app specified by SetKioskAppUserModelId. Only applies when KioskAppUserModelId is set.
     */
    @SerializedName(value = "kioskAppDisplayName", alternate = {"KioskAppDisplayName"})
    @Expose
    public String kioskAppDisplayName;

    /**
     * The Kiosk App User Model Id.
     * Specifies the application user model ID of the app to use with assigned access.
     */
    @SerializedName(value = "kioskAppUserModelId", alternate = {"KioskAppUserModelId"})
    @Expose
    public String kioskAppUserModelId;

    /**
     * The Maintenance Start Time.
     * Specifies the daily start time of maintenance hour.
     */
    @SerializedName(value = "maintenanceStartTime", alternate = {"MaintenanceStartTime"})
    @Expose
    public com.microsoft.graph.core.TimeOfDay maintenanceStartTime;


    /**
     * The raw representation of this class
     */
    private JsonObject rawObject;

    /**
     * The serializer
     */
    private ISerializer serializer;

    /**
     * Gets the raw representation of this class
     *
     * @return the raw representation of this class
     */
    public JsonObject getRawObject() {
        return rawObject;
    }

    /**
     * Gets serializer
     *
     * @return the serializer
     */
    protected ISerializer getSerializer() {
        return serializer;
    }

    /**
     * Sets the raw JSON object
     *
     * @param serializer the serializer
     * @param json the JSON object to set this object to
     */
    public void setRawObject(final ISerializer serializer, final JsonObject json) {
        this.serializer = serializer;
        rawObject = json;

    }
}
