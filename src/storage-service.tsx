import SettingsDocument from "./settingsDocument";
import * as SDK from "azure-devops-extension-sdk";
import { CommonServiceIds, IExtensionDataService, IProjectPageService } from "azure-devops-extension-api";
import { Constants } from "./constants";

const CollectionName: string = "ProjectSettings";
const ScopeType: string = "Default";
const UserPreferencesCollection: string = "UserPreferences";
const UserScopeType: string = "User";

export class StorageService {
    public static Foo: string = "Test";
    private dataService?: IExtensionDataService;
    private _projectId?: string;

    private async getDataService(): Promise<IExtensionDataService> {
        if (this.dataService === undefined) {
            this.dataService = await SDK.getService<IExtensionDataService>('ms.vss-features.extension-data-service');
        }

        if (this._projectId === undefined) {
            const projectService = await SDK.getService<IProjectPageService>(CommonServiceIds.ProjectPageService);
            const project = await projectService.getProject()

            if (project === undefined) {
                throw new Error('Failed to find project');
            }

            this._projectId = project.id;
        }

        return this.dataService;
    }

    public async getSettings(): Promise<SettingsDocument> {
        const dataService = await this.getDataService();
        const dataManager = await dataService.getExtensionDataManager(
            SDK.getExtensionContext().id,
            await SDK.getAccessToken()
        );

        let settingsDocument: SettingsDocument = {
            ...Constants.DefaultSettingsDocument,
            id: this._projectId!
        };
        try {
            const settingsDocumentData = await dataManager.getDocument(CollectionName, this._projectId!, { scopeType: ScopeType });

            settingsDocument = {
                ...Constants.DefaultSettingsDocument,
                ...settingsDocumentData
            };
        } catch (error) {
            settingsDocument = await this.setSettings(settingsDocument);
        }

        return settingsDocument;
    }

    public async setSettings(settingsDocument: SettingsDocument): Promise<SettingsDocument> {
        const dataService = await this.getDataService();
        const dataManager = await dataService.getExtensionDataManager(
            SDK.getExtensionContext().id,
            await SDK.getAccessToken()
        );

        return dataManager.setDocument(CollectionName, settingsDocument, { scopeType: ScopeType });
    }

    public async getLastUsedRepositoryId(): Promise<string | undefined> {
        try {
            const dataService = await this.getDataService();
            const dataManager = await dataService.getExtensionDataManager(
                SDK.getExtensionContext().id,
                await SDK.getAccessToken()
            );
            const doc = await dataManager.getDocument(UserPreferencesCollection, "preferences", { scopeType: UserScopeType });
            return doc?.lastUsedRepositoryId;
        } catch {
            return undefined;
        }
    }

    public async saveLastUsedRepositoryId(repositoryId: string): Promise<void> {
        const dataService = await this.getDataService();
        const dataManager = await dataService.getExtensionDataManager(
            SDK.getExtensionContext().id,
            await SDK.getAccessToken()
        );
        let doc: { id: string; lastUsedRepositoryId: string; __etag?: string } = { id: "preferences", lastUsedRepositoryId: repositoryId };
        try {
            const existing = await dataManager.getDocument(UserPreferencesCollection, "preferences", { scopeType: UserScopeType });
            doc.__etag = existing.__etag;
        } catch {
            // Document doesn't exist yet, create without __etag
        }
        await dataManager.setDocument(UserPreferencesCollection, doc, { scopeType: UserScopeType });
    }
}