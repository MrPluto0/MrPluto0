package errors

type Code int32

const (
	CodeSuccess Code = 0

	CodeUnauthorized                 Code = 991000
	CodeForbidden                    Code = 991001
	CodeServiceAccountNotInWhitelist Code = 991002

	CodeBadRequest         Code = 991100
	CodeInvalidSource      Code = 991101
	CodeInvalidParameter   Code = 991102
	CodeUserConfigNotFound Code = 991103

	CodeChatSessionNotFound     Code = 991200
	CodeChatSessionNotRunning   Code = 991201
	CodeChatSessionNotCreating  Code = 991202
	CodeChatSessionForbidden    Code = 991203
	CodeChatSessionNotRevertable Code = 991204
	CodeChatSessionNotDeletable  Code = 991205

	CodeMessageNotFound Code = 991300

	CodeGitRepoNotFound   Code = 991400
	CodeGitBranchNotFound Code = 991401
	CodeGitTokenNotFound  Code = 991402
	CodeGitTokenInvalid   Code = 991403

	CodeSandboxAllocationInProgress Code = 991500
	CodeUnsupportedSandboxType      Code = 991501
	CodeSoloAgentParallelLimit      Code = 991502

	CodeWebhookNotFound  Code = 991600
	CodeWebhookForbidden Code = 991601

	CodeSSHKeyNotFound      Code = 991700
	CodeSSHKeyAlreadyExists Code = 991701
	CodeInvalidKeyType      Code = 991702

	CodeUserRuleNotFound         Code = 991800
	CodeUserRuleForbidden        Code = 991801
	CodeUserRuleFilenameConflict Code = 991802

	CodeSlashCommandNotFound     Code = 991900
	CodeSlashCommandForbidden    Code = 991901
	CodeSlashCommandNameConflict Code = 991902

	CodeSkillNotFound     Code = 992000
	CodeSkillForbidden    Code = 992001
	CodeSkillNameConflict Code = 992002

	CodeEnvironmentNotFound      Code = 992100
	CodeEnvironmentNameDuplicate Code = 992101
	CodeEnvironmentForbidden     Code = 992102
	CodeExplorerURLNotConfigured Code = 992103

	CodeMCPConfigNotFound       Code = 992200
	CodeMCPConfigForbidden      Code = 992201
	CodeMCPConfigNameConflict   Code = 992202
	CodeMCPConfigSecretsDirty   Code = 992203
	CodeMCPConfigValidationFail Code = 992204

	CodeVersionSnapshotNotFound Code = 992300
	CodeVersionTypeMismatch     Code = 992301

	CodeResourceLimitExceeded Code = 992350

	CodeFileConvertEmptyContent Code = 992400
	CodeFileConvertInvalidType  Code = 992401
	CodeFileConvertFailed       Code = 992402
	CodeFileConvertTimeout      Code = 992403
	CodeFileConvertNotFound     Code = 992404
	CodeFileConvertAccessDenied Code = 992405

	CodeProjectNotFound  Code = 992500
	CodeProjectForbidden Code = 992501

	CodeInternalError       Code = 995000
	CodeServiceUnavailable  Code = 995001
	CodeTimeout             Code = 995002
	CodeDatabaseError       Code = 995003
	CodeExternalService     Code = 995004
	CodeK8sError            Code = 995005
	CodeTCCError            Code = 995006
	CodeCloudIDEError       Code = 995007
)

var codeMessages = map[Code]string{
	CodeSuccess: "success",

	CodeUnauthorized:                 "unauthorized",
	CodeForbidden:                    "forbidden",
	CodeServiceAccountNotInWhitelist: "service account not in whitelist",

	CodeBadRequest:         "bad request",
	CodeInvalidSource:      "invalid source",
	CodeInvalidParameter:   "invalid parameter",
	CodeUserConfigNotFound: "user configuration not found",

	CodeChatSessionNotFound:      "chat session not found",
	CodeChatSessionNotRunning:    "chat session is not running",
	CodeChatSessionNotCreating:   "chat session is not in creating status",
	CodeChatSessionForbidden:     "access to chat session is forbidden",
	CodeChatSessionNotRevertable: "chat session status does not allow revert",
	CodeChatSessionNotDeletable:  "chat session status does not allow delete",

	CodeMessageNotFound: "message not found",

	CodeGitRepoNotFound:   "git repository not found",
	CodeGitBranchNotFound: "git branch not found",
	CodeGitTokenNotFound:  "git token not found",
	CodeGitTokenInvalid:   "git token is invalid",

	CodeSandboxAllocationInProgress: "sandbox allocation is in progress",
	CodeUnsupportedSandboxType:      "unsupported sandbox type",
	CodeSoloAgentParallelLimit:      "solo agent parallel limit reached",

	CodeWebhookNotFound:  "webhook not found",
	CodeWebhookForbidden: "access to webhook is forbidden",

	CodeSSHKeyNotFound:      "ssh key not found",
	CodeSSHKeyAlreadyExists: "ssh key already exists",
	CodeInvalidKeyType:      "invalid key type",

	CodeUserRuleNotFound:         "user rule not found",
	CodeUserRuleForbidden:        "access to user rule is forbidden",
	CodeUserRuleFilenameConflict: "user rule filename already exists",

	CodeSlashCommandNotFound:     "slash command not found",
	CodeSlashCommandForbidden:    "access to slash command is forbidden",
	CodeSlashCommandNameConflict: "slash command name already exists",

	CodeSkillNotFound:     "skill not found",
	CodeSkillForbidden:    "access to skill is forbidden",
	CodeSkillNameConflict: "skill name already exists",

	CodeEnvironmentNotFound:      "environment not found",
	CodeEnvironmentNameDuplicate: "environment name already exists",
	CodeEnvironmentForbidden:     "access to environment is forbidden",
	CodeExplorerURLNotConfigured: "explorer URL is not configured",

	CodeMCPConfigNotFound:       "mcp config not found",
	CodeMCPConfigForbidden:      "access to mcp config is forbidden",
	CodeMCPConfigNameConflict:   "mcp config name already exists",
	CodeMCPConfigSecretsDirty:   "operation failed and secrets may be dirty, recommend deleting and recreating",
	CodeMCPConfigValidationFail: "mcp config validation failed",

	CodeVersionSnapshotNotFound: "version snapshot not found",
	CodeVersionTypeMismatch:     "version type mismatch",

	CodeFileConvertEmptyContent: "file content is empty",
	CodeFileConvertInvalidType:  "invalid target type",
	CodeFileConvertFailed:       "file convert failed",
	CodeFileConvertTimeout:      "file convert timeout",
	CodeFileConvertNotFound:     "file convert not found",
	CodeFileConvertAccessDenied: "access denied to file convert",

	CodeProjectNotFound:  "project not found",
	CodeProjectForbidden: "access to project is forbidden",

	CodeResourceLimitExceeded: "resource limit exceeded",

	CodeInternalError:      "internal server error",
	CodeServiceUnavailable: "service unavailable",
	CodeTimeout:            "request timeout",
	CodeDatabaseError:      "database error",
	CodeExternalService:    "external service error",
	CodeK8sError:           "kubernetes error",
	CodeTCCError:           "tcc service error",
	CodeCloudIDEError:      "cloudide service error",
}

func (c Code) Int32() int32 {
	return int32(c)
}

func (c Code) Message() string {
	if msg, ok := codeMessages[c]; ok {
		return msg
	}
	return "unknown error"
}

func (c Code) WithMessage(msg string) (int32, string) {
	return int32(c), msg
}
