from Crypto.Cipher import AES
from Crypto.Random import get_random_bytes
import json

class EncryptedAccountManager:
    def __init__(self):
        self.key = b'u\xfeT\xcc\x00\x9dZX\x11\x8a\xed/\xb63\xcb4'
        self.accounts = {}
        self.file_path = 'operation\\accounts.bin'
        
    @staticmethod
    def pad(data):
        block_size = AES.block_size
        padding = block_size - len(data) % block_size
        return data + bytes([padding] * padding)

    @staticmethod
    def unpad(data):
        padding = data[-1]
        return data[:-padding]

    def encrypt_aes(self, data):
        iv = get_random_bytes(AES.block_size)
        cipher = AES.new(self.key, AES.MODE_CBC, iv)
        encrypted_data = cipher.encrypt(self.pad(data))
        return iv + encrypted_data

    def decrypt_aes(self, data):
        iv = data[:AES.block_size]
        cipher = AES.new(self.key, AES.MODE_CBC, iv)
        decrypted_data = self.unpad(cipher.decrypt(data[AES.block_size:]))
        return decrypted_data
    
    def add_account(self, username, account_info):
        self.accounts[username] = account_info

    def encrypt_and_save(self):
        encrypted_accounts = self.encrypt_aes(json.dumps(self.accounts).encode('utf-8'))
        
        with open(self.file_path, 'wb') as file:
            file.write(encrypted_accounts)

    def load_and_decrypt(self):
        with open(self.file_path, 'rb') as file:
            encrypted_accounts = file.read()

        decrypted_data = self.decrypt_aes(encrypted_accounts)
        self.accounts = json.loads(decrypted_data.decode('utf-8'))

    def get_account_by_name(self, username):
        return self.accounts.get(username)
    
    def get_all_account_names(self):
        return list(self.accounts.keys())

    def update_account_by_name(self, username, updated_account_info):
        self.accounts[username] = updated_account_info
        
    def delete_account_by_name(self, username):
        if username in self.accounts:
            del self.accounts[username]