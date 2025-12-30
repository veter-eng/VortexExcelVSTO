namespace VortexExcelAddIn.Domain.Interfaces
{
    /// <summary>
    /// Interface para criptografia de credenciais.
    /// Implementa SRP (Single Responsibility Principle) - responsabilidade única de criptografar/descriptografar.
    /// Implementa DIP (Dependency Inversion Principle) - permite diferentes implementações de criptografia.
    /// </summary>
    public interface ICredentialEncryptor
    {
        /// <summary>
        /// Criptografa um texto plano (credencial).
        /// </summary>
        /// <param name="plainText">Texto em formato plano</param>
        /// <returns>Texto criptografado</returns>
        string Encrypt(string plainText);

        /// <summary>
        /// Descriptografa um texto criptografado.
        /// </summary>
        /// <param name="encryptedText">Texto criptografado</param>
        /// <returns>Texto em formato plano</returns>
        string Decrypt(string encryptedText);

        /// <summary>
        /// Verifica se um texto está criptografado.
        /// </summary>
        /// <param name="text">Texto a verificar</param>
        /// <returns>True se o texto está criptografado, False caso contrário</returns>
        bool IsEncrypted(string text);
    }
}
