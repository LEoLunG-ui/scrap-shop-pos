// ThaiIdReader — โปรแกรมช่วยอ่านบัตรประชาชนไทยผ่านเครื่องอ่าน Smart Card (PC/SC)
//
// เขียนด้วย C#/.NET เรียก winscard.dll ของ Windows โดยตรง (ไม่ต้องคอมไพล์ native module
// แบบ Node.js เพราะเครื่องพัฒนา/เครื่องร้านอาจไม่มี Visual Studio C++ Build Tools ติดตั้งไว้)
//
// ชุดคำสั่ง APDU และลอจิกไล่ดึงข้อมูล (GET RESPONSE chaining) ในไฟล์นี้ อ้างอิงตรงจาก
// แอป "ThaiIDCardReader" (nfc-pcsc + native pcsclite) ที่ผู้ใช้ทดสอบแล้วว่าอ่านบัตรผ่านได้จริง
// กับเครื่องอ่าน/บัตรชุดนี้ — จุดสำคัญที่โค้ดรุ่นก่อนหน้าขาดไปคือ: การ์ดตอบกลับ SW1=0x61
// (แปลว่า "มีข้อมูลอีก N ไบต์ รอเรียกด้วย GET RESPONSE") ต้องส่งคำสั่ง GET RESPONSE
// (00 C0 00 00 N) ตามไปอีกต่อจนกว่าจะได้ SW1SW2=9000 ไม่ใช่แค่ส่ง READ BINARY แล้วจบเลย
// ถ้าข้ามขั้นตอนนี้ไป การ์ดจะค้างสถานะ "รอ GET RESPONSE" ทำให้คำสั่งถัดไปพังหมด
// (คือสาเหตุที่โค้ดรุ่นก่อนอ่าน SELECT ได้ (เห็น 61 0A ใน debug log) แต่คำสั่งอ่านข้อมูล
// ถัดไปทุกคำสั่งพังด้วย SCardTransmit ล้มเหลว/การสื่อสารขาดหายตลอด)
//
// ทำงานแบบ "watch mode" — วนเช็คบัตรตลอดเวลา พิมพ์ผลลัพธ์เป็น JSON บรรทัดละ 1 ก้อน (NDJSON)
// ออกทาง stdout ให้โปรเซสหลัก (main.js) อ่านและส่งต่อให้หน้าจอ
//
// สำคัญ: ให้ "โปรแกรม POS หลัก" เป็นตัวเดียวที่รันไฟล์นี้เท่านั้น (ห้ามมีสองโปรเซสรันพร้อมกัน
// เคยเจอปัญหาสองโปรเซสแย่งกันคุยกับเครื่องอ่านจนอ่านบัตรไม่ออกเลยทั้งคู่มาแล้วจริง) โปรแกรม
// อื่นที่ต้องการข้อมูลบัตร ให้รับผ่าน WebSocket ของโปรแกรม POS หลักแทน ไม่ต้องรันไฟล์นี้เอง

using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Threading;

namespace ThaiIdReader;

class CardData
{
    public string? Cid { get; set; }
    public string? ThaiName { get; set; }
    public string? EnglishName { get; set; }
    public string? BirthDateIso { get; set; }
    public string? Gender { get; set; }
    public string? Address { get; set; }
    public string? IssueDateIso { get; set; }
    public string? ExpireDateIso { get; set; }
    public string? PhotoBase64Jpeg { get; set; }
}

[StructLayout(LayoutKind.Sequential)]
struct ScardIoRequest
{
    public uint dwProtocol;
    public uint cbPciLength;
}

static class Native
{
    [DllImport("winscard.dll")]
    internal static extern int SCardEstablishContext(uint dwScope, IntPtr notUsed1, IntPtr notUsed2, out IntPtr phContext);

    [DllImport("winscard.dll")]
    internal static extern int SCardReleaseContext(IntPtr hContext);

    [DllImport("winscard.dll", CharSet = CharSet.Ansi)]
    internal static extern int SCardListReaders(IntPtr hContext, byte[]? mszGroups, byte[]? mszReaders, ref int pcchReaders);

    [DllImport("winscard.dll", CharSet = CharSet.Ansi)]
    internal static extern int SCardConnect(IntPtr hContext, string szReader, uint dwShareMode, uint dwPreferredProtocols, out IntPtr phCard, out uint pdwActiveProtocol);

    [DllImport("winscard.dll")]
    internal static extern int SCardDisconnect(IntPtr hCard, uint dwDisposition);

    [DllImport("winscard.dll")]
    internal static extern int SCardTransmit(IntPtr hCard, ref ScardIoRequest pioSendPci, byte[] pbSendBuffer, int cbSendLength, IntPtr pioRecvPci, byte[] pbRecvBuffer, ref int pcbRecvLength);
}

// คำสั่ง APDU ที่ยืนยันแล้วว่าใช้งานได้จริง (คัดลอกมาจากแอปที่ทดสอบผ่าน) — อย่าแก้ตัวเลข
// offset/length พวกนี้เอง ถ้าจะปรับให้อ้างอิงจากแหล่งเดียวกันเท่านั้น
static class Apdu
{
    public static readonly byte[] SelectApplet = { 0x00, 0xA4, 0x04, 0x00, 0x08, 0xA0, 0x00, 0x00, 0x00, 0x54, 0x48, 0x00, 0x01 };

    public static readonly byte[] Cid = { 0x80, 0xB0, 0x00, 0x04, 0x02, 0x00, 0x0D };
    public static readonly byte[] ThFullname = { 0x80, 0xB0, 0x00, 0x11, 0x02, 0x00, 0x64 };
    public static readonly byte[] EnFullname = { 0x80, 0xB0, 0x00, 0x75, 0x02, 0x00, 0x64 };
    public static readonly byte[] Birthdate = { 0x80, 0xB0, 0x00, 0xD9, 0x02, 0x00, 0x08 };
    public static readonly byte[] Gender = { 0x80, 0xB0, 0x00, 0xE1, 0x02, 0x00, 0x01 };
    public static readonly byte[] IssueDate = { 0x80, 0xB0, 0x01, 0x67, 0x02, 0x00, 0x08 };
    public static readonly byte[] ExpireDate = { 0x80, 0xB0, 0x01, 0x6F, 0x02, 0x00, 0x08 };
    public static readonly byte[] Address = { 0x80, 0xB0, 0x15, 0x79, 0x02, 0x00, 0x64 };

    // 20 บล็อกความยาวคงที่ (P1,P2) ที่ต่อกันเป็นรูปถ่าย JPEG บนชิป
    public static readonly (byte p1, byte p2)[] PhotoBlocks =
    {
        (0x01,0x7B),(0x02,0x7A),(0x03,0x79),(0x04,0x78),(0x05,0x77),
        (0x06,0x76),(0x07,0x75),(0x08,0x74),(0x09,0x73),(0x0A,0x72),
        (0x0B,0x71),(0x0C,0x70),(0x0D,0x6F),(0x0E,0x6E),(0x0F,0x6D),
        (0x10,0x6C),(0x11,0x6B),(0x12,0x6A),(0x13,0x69),(0x14,0x68),
    };

    public static byte[] PhotoBlockApdu(byte p1, byte p2) => new byte[] { 0x80, 0xB0, p1, p2, 0x02, 0x00, 0xFF };
}

class Program
{
    const uint SCARD_SCOPE_USER = 0;
    const uint SCARD_SHARE_SHARED = 2;
    const uint SCARD_PROTOCOL_T0 = 1;
    const uint SCARD_PROTOCOL_T1 = 2;
    const uint SCARD_LEAVE_CARD = 0;
    const int SCARD_S_SUCCESS = 0;

    static bool debug = false;

    static int Main(string[] args)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        debug = args.Length > 0 && args[0] == "--debug";

        WriteEvent(new { type = "starting" });

        int estRc = Native.SCardEstablishContext(SCARD_SCOPE_USER, IntPtr.Zero, IntPtr.Zero, out IntPtr ctx);
        if (estRc != SCARD_S_SUCCESS)
        {
            WriteEvent(new { type = "fatal_error", message = "เชื่อมต่อ PC/SC ไม่สำเร็จ (SCardEstablishContext)", code = estRc });
            return 1;
        }

        bool cardPresentLastLoop = false;
        string? lastCid = null;

        while (true)
        {
            string? readerName = GetFirstReader(ctx);
            if (readerName == null)
            {
                if (cardPresentLastLoop) { WriteEvent(new { type = "card_removed" }); cardPresentLastLoop = false; lastCid = null; }
                WriteEvent(new { type = "no_reader" });
                Thread.Sleep(2500);
                continue;
            }

            int connRc = Native.SCardConnect(ctx, readerName, SCARD_SHARE_SHARED, SCARD_PROTOCOL_T0 | SCARD_PROTOCOL_T1, out IntPtr hCard, out uint activeProtocol);
            if (connRc != SCARD_S_SUCCESS)
            {
                if (cardPresentLastLoop) { WriteEvent(new { type = "card_removed" }); cardPresentLastLoop = false; lastCid = null; }
                Thread.Sleep(700);
                continue;
            }

            try
            {
                var pci = new ScardIoRequest { dwProtocol = activeProtocol, cbPciLength = (uint)Marshal.SizeOf<ScardIoRequest>() };
                var result = ReadThaiIdCardWithRetry(hCard, pci);
                if (result.Cid != null && result.Cid != lastCid)
                {
                    WriteEvent(new { type = "card_data", data = result });
                    lastCid = result.Cid;
                }
                cardPresentLastLoop = true;
            }
            catch (Exception ex)
            {
                WriteEvent(new { type = "read_error", message = ex.Message });
            }
            finally
            {
                Native.SCardDisconnect(hCard, SCARD_LEAVE_CARD);
            }

            Thread.Sleep(1200);
        }
    }

    static string? GetFirstReader(IntPtr hContext)
    {
        int len = 0;
        int rc = Native.SCardListReaders(hContext, null, null, ref len);
        if (rc != SCARD_S_SUCCESS || len <= 0) return null;
        var buf = new byte[len];
        rc = Native.SCardListReaders(hContext, null, buf, ref len);
        if (rc != SCARD_S_SUCCESS) return null;
        string all = Encoding.ASCII.GetString(buf);
        string first = all.Split('\0')[0];
        return string.IsNullOrEmpty(first) ? null : first;
    }

    // ส่ง APDU ดิบหนึ่งคำสั่ง คืนค่าทั้งหมดที่การ์ดตอบกลับ (รวม SW1 SW2 สองไบต์ท้าย)
    static byte[] RawTransmit(IntPtr hCard, ref ScardIoRequest pci, byte[] apdu)
    {
        var recv = new byte[512];
        int recvLen = recv.Length;
        int rc = Native.SCardTransmit(hCard, ref pci, apdu, apdu.Length, IntPtr.Zero, recv, ref recvLen);
        if (debug)
        {
            Console.Error.WriteLine($"[apdu] send={BitConverter.ToString(apdu)} rc={rc} recv={BitConverter.ToString(recv, 0, Math.Max(recvLen, 0))}");
        }
        if (rc != SCARD_S_SUCCESS) throw new Exception($"SCardTransmit ล้มเหลว (rc={rc})");
        if (recvLen < 2) throw new Exception("การ์ดตอบกลับสั้นผิดปกติ");
        Array.Resize(ref recv, recvLen);
        return recv;
    }

    // ไล่ดึงข้อมูลตามมาตรฐาน ISO 7816-4: ถ้าการ์ดตอบ SW1=0x61 (ยังมีข้อมูลอีก N ไบต์) ต้องส่ง
    // GET RESPONSE (00 C0 00 00 N) ตามไปเรื่อยๆ จนกว่าจะได้ SW1SW2=9000 — ถ้า SW1=0x6C คือ
    // Le ที่ส่งไปผิด ต้องส่งคำสั่งเดิมซ้ำโดยใช้ Le ใหม่ที่การ์ดบอกมา (sw2)
    static byte[] Transmit(IntPtr hCard, ref ScardIoRequest pci, byte[] apdu)
    {
        var res = RawTransmit(hCard, ref pci, apdu);
        byte sw1 = res[res.Length - 2];
        byte sw2 = res[res.Length - 1];
        var data = new byte[res.Length - 2];
        Array.Copy(res, data, data.Length);

        while (sw1 == 0x61)
        {
            var getResponse = new byte[] { 0x00, 0xC0, 0x00, 0x00, sw2 };
            res = RawTransmit(hCard, ref pci, getResponse);
            sw1 = res[res.Length - 2];
            sw2 = res[res.Length - 1];
            var chunk = new byte[res.Length - 2];
            Array.Copy(res, chunk, chunk.Length);
            var merged = new byte[data.Length + chunk.Length];
            Array.Copy(data, merged, data.Length);
            Array.Copy(chunk, 0, merged, data.Length, chunk.Length);
            data = merged;
        }

        if (sw1 == 0x6C)
        {
            var retryApdu = new byte[5];
            Array.Copy(apdu, retryApdu, 4);
            retryApdu[4] = sw2;
            return Transmit(hCard, ref pci, retryApdu);
        }

        if (!(sw1 == 0x90 && sw2 == 0x00))
        {
            throw new Exception($"การ์ดตอบกลับสถานะผิดพลาด {sw1:x2}{sw2:x2} สำหรับคำสั่ง {BitConverter.ToString(apdu)}");
        }

        return data;
    }

    static string ReadTextField(IntPtr hCard, ref ScardIoRequest pci, byte[] apdu)
    {
        var raw = Transmit(hCard, ref pci, apdu);
        var thaiEnc = Encoding.GetEncoding(874); // เทียบเท่า TIS-620 ในช่วงตัวอักษรไทย
        return thaiEnc.GetString(raw).Replace("\0", "").Trim();
    }

    static string CleanName(string raw)
    {
        // ฟิลด์ชื่อในบัตรคั่นด้วย # และเติมช่องว่างท้ายให้เต็มความยาวคงที่
        var parts = raw.Split('#');
        var cleaned = new List<string>();
        foreach (var p in parts)
        {
            var t = p.Trim();
            if (t.Length > 0) cleaned.Add(t);
        }
        return string.Join(" ", cleaned);
    }

    static string GenderLabel(string raw)
    {
        var t = raw.Trim();
        if (t == "1") return "ชาย";
        if (t == "2") return "หญิง";
        return t;
    }

    static string? BuddhistDateToIso(string raw)
    {
        var digits = raw.Trim();
        if (digits.Length != 8 || digits == "00000000" || !long.TryParse(digits, out _)) return null;
        int beYear = int.Parse(digits.Substring(0, 4));
        int month = int.Parse(digits.Substring(4, 2));
        int day = int.Parse(digits.Substring(6, 2));
        int ceYear = beYear - 543;
        try
        {
            var d = new DateTime(ceYear, month, day);
            return d.ToString("yyyy-MM-dd");
        }
        catch { return null; }
    }

    static string FormatAddress(string raw)
    {
        var fields = raw.Split('#');
        for (int i = 0; i < fields.Length; i++) fields[i] = fields[i].Trim();
        string Get(int i) => i < fields.Length ? fields[i] : "";
        var houseNo = Get(0); var moo = Get(1); var trok = Get(2); var soi = Get(3);
        var road = Get(4); var tambon = Get(5); var amphoe = Get(6); var province = Get(7);

        var pieces = new List<string>();
        if (houseNo.Length > 0) pieces.Add($"บ้านเลขที่ {houseNo}");
        if (moo.Length > 0) pieces.Add($"หมู่ที่ {moo}");
        if (trok.Length > 0) pieces.Add($"ตรอก{trok}");
        if (soi.Length > 0) pieces.Add($"ซอย{soi}");
        if (road.Length > 0) pieces.Add($"ถนน{road}");
        if (tambon.Length > 0) pieces.Add($"ต.{tambon}");
        if (amphoe.Length > 0) pieces.Add($"อ.{amphoe}");
        if (province.Length > 0) pieces.Add($"จ.{province}");
        return string.Join(" ", pieces);
    }

    // เครื่องอ่านบางรุ่นเจอ error ชั่วคราวในรอบแรกทันทีหลังเสียบบัตรใหม่ (handle ค้าง/การ์ดยัง
    // ตั้งตัวไม่เสร็จ) — ลองต่อการ์ดใหม่แล้วอ่านซ้ำอีกครั้งก่อนถือว่าล้มเหลวจริง
    static CardData ReadThaiIdCardWithRetry(IntPtr hCard, ScardIoRequest pci, int attempt = 1)
    {
        try
        {
            return ReadThaiIdCard(hCard, pci);
        }
        catch when (attempt < 2)
        {
            Thread.Sleep(250);
            return ReadThaiIdCardWithRetry(hCard, pci, attempt + 1);
        }
    }

    static CardData ReadThaiIdCard(IntPtr hCard, ScardIoRequest pciIn)
    {
        var pci = pciIn;
        // ให้เวลาการ์ดตั้งตัวสักครู่หลังต่อสำเร็จ ก่อนส่งคำสั่งแรก (เครื่องอ่านบางรุ่นปฏิเสธ
        // คำสั่งที่ส่งทันทีหลัง connect)
        Thread.Sleep(150);

        Transmit(hCard, ref pci, Apdu.SelectApplet);

        string cid = ReadTextField(hCard, ref pci, Apdu.Cid).Replace(" ", "");
        string thName = CleanName(ReadTextField(hCard, ref pci, Apdu.ThFullname));
        string enName = CleanName(ReadTextField(hCard, ref pci, Apdu.EnFullname));
        string dobRaw = ReadTextField(hCard, ref pci, Apdu.Birthdate);
        string genderRaw = ReadTextField(hCard, ref pci, Apdu.Gender);
        string issueRaw = ReadTextField(hCard, ref pci, Apdu.IssueDate);
        string expireRaw = ReadTextField(hCard, ref pci, Apdu.ExpireDate);
        string addressRaw = ReadTextField(hCard, ref pci, Apdu.Address);

        string? photoBase64 = null;
        try
        {
            using var photoStream = new MemoryStream();
            foreach (var (p1, p2) in Apdu.PhotoBlocks)
            {
                var chunk = Transmit(hCard, ref pci, Apdu.PhotoBlockApdu(p1, p2));
                photoStream.Write(chunk, 0, chunk.Length);
            }
            photoBase64 = Convert.ToBase64String(photoStream.ToArray());
        }
        catch (Exception ex)
        {
            if (debug) Console.Error.WriteLine($"[photo] อ่านรูปไม่สำเร็จ: {ex.Message}");
        }

        return new CardData
        {
            Cid = cid,
            ThaiName = thName,
            EnglishName = enName,
            BirthDateIso = BuddhistDateToIso(dobRaw),
            Gender = GenderLabel(genderRaw),
            Address = FormatAddress(addressRaw),
            IssueDateIso = BuddhistDateToIso(issueRaw),
            ExpireDateIso = BuddhistDateToIso(expireRaw),
            PhotoBase64Jpeg = photoBase64,
        };
    }

    // ต้องใช้ camelCase (cid, thaiName, photoBase64Jpeg, ...) เพราะฝั่งหน้าจอ (renderer/index.html,
    // card-reader-app, card-reader-test) ทุกตัวคาดหวังชื่อฟิลด์แบบนี้ทั้งหมด — ค่าเริ่มต้นของ
    // JsonSerializer จะสะกดตามชื่อ property C# (PascalCase: Cid, ThaiName) ซึ่งไม่ตรงกัน
    static readonly JsonSerializerOptions JsonOpts = new() { PropertyNamingPolicy = JsonNamingPolicy.CamelCase };

    static void WriteEvent(object obj)
    {
        Console.WriteLine(JsonSerializer.Serialize(obj, JsonOpts));
        Console.Out.Flush();
    }
}
