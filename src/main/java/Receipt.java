import lombok.Getter;
import lombok.RequiredArgsConstructor;

@RequiredArgsConstructor
@Getter
public class Receipt {
    private final String number; // 1 通し番号
    private final String issueDate; // 2 発行日
    private final String name; // 3. 宛先
    private final String reissue; // 4. 再発行
    private final String price; // 5. 金額

    // 6. 上記の金額文言?
    // 7. 但し文言?

    private final String site; // 8.サイト名
    private final String orderNumber; // 9. 注文番号
    private final String deliveryNumber; // 10.出荷番号
    @Direction(column = "date2")
    private final String receiptDate; // 11.領収日
    private final String note; // 12. 備考

    private final String receiptBy; // 13. 領収書発行者名
    private final String msp; // 14. MSPオンラインショップ名
    private final String customer; // 15. カスタマーセンター連絡先
    private final String tel; // 電話番号

}
