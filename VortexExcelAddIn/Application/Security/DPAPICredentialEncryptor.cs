using System;
using System.Security.Cryptography;
using System.Text;
using VortexExcelAddIn.Domain.Interfaces;
using VortexExcelAddIn.Services;

namespace VortexExcelAddIn.Application.Security
{
    /// <summary>
    /// Implementação de criptografia de credenciais usando Windows DPAPI (Data Protection API).
    /// Implementa SRP (Single Responsibility Principle) - responsabilidade única de criptografar/descriptografar.
    /// </summary>
    public class DPAPICredentialEncryptor : ICredentialEncryptor
    {
        private const string ENCRYPTION_PREFIX = "DPAPI:";

        /// <summary>
        /// Criptografa um texto plano usando DPAPI com escopo de usuário atual.
        /// </summary>
        /// <param name="plainText">Texto em formato plano</param>
        /// <returns>Texto criptografado com prefixo "DPAPI:"</returns>
        public string Encrypt(string plainText)
        {
            if (string.IsNullOrEmpty(plainText))
                return plainText;

            // Evita criptografar múltiplas vezes
            if (IsEncrypted(plainText))
            {
                LoggingService.Debug("Texto já está criptografado, retornando sem modificação");
                return plainText;
            }

            try
            {
                byte[] plainBytes = Encoding.UTF8.GetBytes(plainText);

                // DataProtectionScope.CurrentUser - apenas o usuário atual pode descriptografar
                // Isso significa que a credencial não funcionará em outra máquina ou outro usuário
                byte[] encryptedBytes = ProtectedData.Protect(
                    plainBytes,
                    null, // entropy opcional (não usado para simplicidade)
                    DataProtectionScope.CurrentUser
                );

                string base64 = Convert.ToBase64String(encryptedBytes);
                string result = ENCRYPTION_PREFIX + base64;

                LoggingService.Debug("Credencial criptografada com sucesso usando DPAPI");
                return result;
            }
            catch (CryptographicException ex)
            {
                LoggingService.Error("Erro ao criptografar credencial com DPAPI", ex);
                throw new CryptographicException("Falha ao criptografar credencial. Verifique as permissões do sistema.", ex);
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro inesperado ao criptografar credencial", ex);
                throw;
            }
        }

        /// <summary>
        /// Descriptografa um texto criptografado usando DPAPI.
        /// </summary>
        /// <param name="encryptedText">Texto criptografado com prefixo "DPAPI:"</param>
        /// <returns>Texto em formato plano</returns>
        public string Decrypt(string encryptedText)
        {
            if (string.IsNullOrEmpty(encryptedText))
                return encryptedText;

            // Se não está criptografado, retornar como está (compatibilidade)
            if (!IsEncrypted(encryptedText))
            {
                LoggingService.Warn("Tentativa de descriptografar texto não criptografado");
                return encryptedText;
            }

            try
            {
                // Remover prefixo "DPAPI:"
                string base64 = encryptedText.Substring(ENCRYPTION_PREFIX.Length);
                byte[] encryptedBytes = Convert.FromBase64String(base64);

                byte[] plainBytes = ProtectedData.Unprotect(
                    encryptedBytes,
                    null,
                    DataProtectionScope.CurrentUser
                );

                string result = Encoding.UTF8.GetString(plainBytes);

                LoggingService.Debug("Credencial descriptografada com sucesso");
                return result;
            }
            catch (CryptographicException ex)
            {
                LoggingService.Error("Erro ao descriptografar credencial com DPAPI", ex);
                throw new CryptographicException(
                    "Falha ao descriptografar credencial. A credencial pode estar corrompida ou foi criada por outro usuário/máquina.",
                    ex);
            }
            catch (FormatException ex)
            {
                LoggingService.Error("Formato inválido de credencial criptografada", ex);
                throw new CryptographicException("Formato de credencial criptografada inválido", ex);
            }
            catch (Exception ex)
            {
                LoggingService.Error("Erro inesperado ao descriptografar credencial", ex);
                throw;
            }
        }

        /// <summary>
        /// Verifica se um texto está criptografado (possui o prefixo "DPAPI:").
        /// </summary>
        /// <param name="text">Texto a verificar</param>
        /// <returns>True se o texto está criptografado, False caso contrário</returns>
        public bool IsEncrypted(string text)
        {
            return !string.IsNullOrEmpty(text) && text.StartsWith(ENCRYPTION_PREFIX);
        }
    }
}
