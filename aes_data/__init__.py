# -*-  coding: utf-8  -*-
__author__ = '10409003'
import hashlib, os
import base64
from Crypto.Cipher import AES

class aes_data:
    global SALT_LENGTH,DERIVATION_ROUNDS,BLOCK_SIZE,KEY_SIZE,MODE

    def __init__(self):
        self.SALT_LENGTH = 16
        self.DERIVATION_ROUNDS=168
        self.BLOCK_SIZE = 16
        self.KEY_SIZE = 16
        self.MODE = AES.MODE_CBC

    def AESencrypt(self,password, plaintext, Needbase64=False):
        salt = os.urandom(self.SALT_LENGTH)
        iv = os.urandom(self.BLOCK_SIZE)
        print self.BLOCK_SIZE
        print iv
        paddingLength = 16 - (len(plaintext) % 16)
        paddedPlaintext = plaintext+chr(paddingLength)*paddingLength
        derivedKey = password
        for i in range(0,self.DERIVATION_ROUNDS):
            derivedKey = hashlib.sha256(derivedKey+salt).digest()
        derivedKey = derivedKey[:self.KEY_SIZE]
        cipherSpec = AES.new(derivedKey, self.MODE, iv)
        ciphertext = cipherSpec.encrypt(paddedPlaintext)
        ciphertext = ciphertext + iv + salt

        if Needbase64:
            return base64.b64encode(ciphertext)
        else:
            return ciphertext.encode("hex")

    def AESdecrypt(self,password, ciphertext, Needbase64=False):
        if Needbase64:
            decodedCiphertext = base64.b64decode(ciphertext)
        else:
            decodedCiphertext = ciphertext.decode("hex")
        startIv = len(decodedCiphertext)-self.BLOCK_SIZE-self.SALT_LENGTH
        startSalt = len(decodedCiphertext)-self.SALT_LENGTH
        data, iv, salt = decodedCiphertext[:startIv], decodedCiphertext[startIv:startSalt], decodedCiphertext[startSalt:]
        derivedKey = password
        for i in range(0, self.DERIVATION_ROUNDS):
            derivedKey = hashlib.sha256(derivedKey+salt).digest()
        derivedKey = derivedKey[:self.KEY_SIZE]
        cipherSpec = AES.new(derivedKey, self.MODE, iv)
        plaintextWithPadding = cipherSpec.decrypt(data)
        paddingLength = ord(plaintextWithPadding[-1])
        plaintext = plaintextWithPadding[:-paddingLength]
        return plaintext

if __name__=="__main__":
    aes=aes_data()
    print aes.AESdecrypt('password','ar4p0l+iAXLUPM3GpkfeuhKbKSGFiNBAHJwt/W5EXwDV1wljzEotzr4K0lcVXQzn',True)