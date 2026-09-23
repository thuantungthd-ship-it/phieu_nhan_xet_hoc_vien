import re
import random
import statistics
import datetime
import openpyxl
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# -*- coding: utf-8 -*-
"""Ngan hang noi dung: nhan xet, kien nghi, loi khuyen theo ky nang, thoi gian hoc toi uu."""

STANDARD = {
 "A+": {"label": "Outstanding – Xuất sắc", "items": [
   "Học viên đạt kết quả xuất sắc, thể hiện sự vượt trội ở tất cả các kỹ năng nghe, nói, đọc, viết. Khả năng vận dụng kiến thức vào giao tiếp thực tế rất linh hoạt và tự nhiên. Học viên luôn chủ động, tích cực và là tấm gương cho các bạn trong lớp noi theo.",
   "Học viên hoàn thành khóa học với thành tích vượt trội, nắm vững toàn diện kiến thức và kỹ năng. Phản xạ giao tiếp nhanh nhạy, tư duy ngôn ngữ tốt, luôn thể hiện tinh thần học tập nghiêm túc và cầu tiến. Đây là kết quả rất đáng khích lệ.",
 ], "kien_nghi": "Học viên cần tiếp tục duy trì và phát huy tối đa kết quả học tập xuất sắc hiện tại, thử sức với các nội dung nâng cao, tham gia các hoạt động ngoại khóa bằng tiếng Anh và đặt mục tiêu chứng chỉ ở cấp độ cao hơn để phát triển toàn diện năng lực ngôn ngữ."},
 "A": {"label": "Excellent – Rất giỏi", "items": [
   "Học viên có kết quả học tập tốt và nắm vững các kiến thức trọng tâm của khóa học. Các kỹ năng nghe, nói, đọc, viết được phát triển tương đối đồng đều; học viên có khả năng vận dụng kiến thức vào giao tiếp và các tình huống thực tế. Học viên có ý thức học tập tốt, tích cực tham gia các hoạt động trên lớp.",
   "Học viên duy trì kết quả học tập tốt và ổn định trong suốt quá trình học. Các kỹ năng nghe, nói, đọc, viết được phát triển tương đối đồng đều. Học viên có ý thức học tập tốt và nên tiếp tục duy trì phong độ, đồng thời mở rộng vốn từ và tăng cường các hoạt động thực hành nâng cao.",
 ], "kien_nghi": "Học viên cần tiếp tục duy trì và phát huy kết quả học tập hiện tại, đồng thời tăng cường các hoạt động thực hành và phát triển kỹ năng nâng cao. Nên chủ động sử dụng tiếng Anh trong giao tiếp, rèn luyện tư duy và phản xạ; phụ huynh có thể theo dõi thêm những kỹ năng có dấu hiệu giảm để hỗ trợ kịp thời."},
 "B+": {"label": "Very good – Giỏi", "items": [
   "Học viên có kết quả học tập tốt và đang có sự tiến bộ rõ rệt. Khả năng tiếp thu, vận dụng kiến thức và thực hành các kỹ năng tiếng Anh ngày càng tốt. Học viên có tinh thần học tập tích cực và nên tiếp tục duy trì phong độ, đồng thời thử sức với các nội dung nâng cao.",
   "Học viên đạt kết quả khá tốt, các kỹ năng nghe, nói, đọc, viết phát triển khá đồng đều. Học viên có tinh thần cầu tiến, sẵn sàng tiếp thu góp ý và cải thiện. Cần tiếp tục rèn luyện thêm để tiến gần hơn tới mức xuất sắc.",
 ], "kien_nghi": "Học viên cần tiếp tục duy trì đà tiến bộ, tăng cường thực hành các kỹ năng còn chưa đồng đều để đạt kết quả cao hơn, luyện tập thêm ở nhà, tham gia các hoạt động giao tiếp thực tế và đặt mục tiêu vươn lên mức xuất sắc trong giai đoạn tiếp theo."},
 "B": {"label": "Good – Khá", "items": [
   "Học viên đạt yêu cầu của khóa học và đã nắm được các kiến thức cơ bản. Các kỹ năng nghe, nói, đọc, viết đang có sự tiến bộ, tuy nhiên khả năng vận dụng kiến thức và phản xạ giao tiếp vẫn cần được cải thiện thêm. Học viên nên duy trì việc học tập đều đặn, tăng cường luyện tập và chủ động tham gia các hoạt động trên lớp để nâng cao kết quả.",
   "Học viên có kết quả học tập khá, nắm được phần lớn kiến thức trọng tâm. Một số kỹ năng còn chưa thật sự vững, cần luyện tập thêm để phản xạ nhanh và tự nhiên hơn trong giao tiếp. Học viên nên duy trì tinh thần học tập tích cực hiện tại.",
 ], "kien_nghi": "Học viên cần tiếp tục duy trì quá trình học tập, tập trung củng cố các kiến thức, kỹ năng còn hạn chế và tăng cường thực hành. Phụ huynh nên đồng hành, khuyến khích học viên nâng cao tính chủ động, ý thức tự học và mức độ tham gia trên lớp."},
 "C+": {"label": "Average – Trung bình", "items": [
   "Học viên duy trì kết quả học tập tương đối ổn định và đáp ứng các yêu cầu cơ bản của khóa học. Các kỹ năng đang được phát triển ở mức phù hợp. Học viên cần tiếp tục duy trì việc học đều đặn và tăng cường luyện tập để nâng cao khả năng vận dụng kiến thức.",
   "Học viên đạt mức yêu cầu cơ bản nhưng vẫn còn một số điểm cần được hỗ trợ thêm. Học viên cần củng cố từ vựng, ngữ pháp và tăng cường thực hành các kỹ năng còn hạn chế. Với sự hướng dẫn thường xuyên và kế hoạch học tập phù hợp, học viên có thể cải thiện kết quả trong thời gian tới.",
 ], "kien_nghi": "Học viên cần tăng cường ôn tập, củng cố kiến thức nền tảng và làm thêm bài tập bổ trợ. Phụ huynh nên khuyến khích học viên chủ động ôn tập, tham gia đầy đủ các hoạt động trên lớp và theo dõi sát tiến độ để không bị tụt lại so với yêu cầu chung."},
 "C": {"label": "Need improvement – Cần cố gắng hơn", "items": [
   "Học viên chưa đạt kết quả như mong đợi và còn hạn chế ở một số kỹ năng tiếng Anh. Khả năng ghi nhớ từ vựng, vận dụng ngữ pháp và phản xạ giao tiếp cần được cải thiện. Học viên cần tăng cường luyện tập, hoàn thành đầy đủ bài tập và chủ động tham gia các hoạt động trên lớp.",
   "Học viên hiện vẫn còn một số hạn chế trong việc tiếp thu và vận dụng kiến thức, tuy nhiên đã có sự tiến bộ so với giai đoạn trước. Khả năng tham gia học tập và thực hành đang được cải thiện. Cần tiếp tục duy trì việc học đều đặn, củng cố kiến thức nền tảng và luyện tập thường xuyên để đạt kết quả tốt hơn.",
 ], "kien_nghi": "Học viên cần củng cố kiến thức nền tảng và tăng cường luyện tập các kỹ năng còn hạn chế. Phụ huynh nên đồng hành, khuyến khích học viên nâng cao tính chủ động trong học tập và duy trì việc ôn tập đều đặn."},
 "D": {"label": "Fail – Không đạt", "items": [
   "Kết quả học tập của học viên có dấu hiệu giảm so với giai đoạn trước. Học viên còn gặp khó khăn trong việc ghi nhớ và vận dụng kiến thức, đồng thời mức độ chủ động trong học tập chưa cao. Cần sớm có kế hoạch ôn tập và theo dõi sát hơn để cải thiện kết quả.",
   "Học viên đang gặp nhiều khó khăn trong quá trình học tập và cần được hỗ trợ thêm. Một số kiến thức nền tảng chưa được nắm vững, ảnh hưởng đến khả năng phát triển các kỹ năng nghe, nói, đọc và viết. Phụ huynh cần quan tâm hỗ trợ thêm, đồng thời xây dựng kế hoạch học tập phù hợp cho học viên.",
 ], "kien_nghi": "Học viên chưa đạt yêu cầu của khóa học và cần được hỗ trợ đặc biệt để củng cố lại kiến thức nền tảng. Phụ huynh nên quan tâm sát sao, xác định nguyên nhân cụ thể và cân nhắc phương án chuyển lớp hoặc học lại để học viên theo kịp chương trình trong giai đoạn tiếp theo."},
}

GRADE_ORDER = ["A+", "A", "B+", "B", "C+", "C", "D"]

def grade_from_total(total):
    if total >= 97.5: return "A+"
    if total >= 94.5: return "A"
    if total >= 89.5: return "B+"
    if total >= 79.5: return "B"
    if total >= 74.5: return "C+"
    if total >= 64.5: return "C"
    return "D"

IELTS_BANDS = [
 (4.0, "<4.0", "Sơ cấp (cần xây lại nền tảng)", [
   "Học viên đạt Band điểm dưới 4.0, tương đương trình độ sơ cấp. Học viên còn gặp nhiều khó khăn trong việc sử dụng tiếng Anh ở cả 4 kỹ năng, vốn từ vựng và ngữ pháp còn hạn chế, cần được xây dựng lại nền tảng một cách bài bản.",
   "Kết quả thi hiện tại của học viên còn cách khá xa mục tiêu Band điểm mong muốn. Khả năng nghe hiểu, đọc hiểu, diễn đạt nói và viết đều cần được củng cố từ gốc. Học viên cần một lộ trình học tập trung vào nền tảng ngữ pháp - từ vựng trước khi luyện đề chuyên sâu.",
 ], "Học viên nên bắt đầu lại từ các khóa nền tảng (Foundation/Elementary) để củng cố ngữ pháp và từ vựng cơ bản trước khi quay lại luyện thi IELTS. Phụ huynh nên đồng hành xây dựng lộ trình học phù hợp, tăng cường số buổi học và ưu tiên các kỹ năng nền tảng."),
 (4.5, "4.0–4.5", "Tiền trung cấp (cần cải thiện)", [
   "Học viên đạt Band 4.0-4.5, ở mức tiền trung cấp. Học viên đã có một số nền tảng tiếng Anh nhất định nhưng khả năng vận dụng vào bài thi thực tế (đặc biệt Nói và Viết) còn hạn chế. Cần tăng cường luyện tập theo dạng đề và mở rộng vốn từ vựng học thuật.",
   "Kết quả hiện tại cho thấy học viên cần cải thiện thêm để đạt mục tiêu Band điểm mong muốn. Một số kỹ năng còn yếu hơn các kỹ năng khác, ảnh hưởng đến điểm tổng. Học viên nên tập trung luyện tập có định hướng theo từng dạng bài thi.",
 ], "Học viên cần tăng cường luyện đề theo từng kỹ năng, đặc biệt các kỹ năng còn yếu, đồng thời mở rộng vốn từ vựng theo chủ đề thường gặp trong IELTS. Nên duy trì lịch học đều đặn và làm quen dần với format đề thi thật."),
 (5.5, "5.0–5.5", "Trung cấp (khá ổn định)", [
   "Học viên đạt Band 5.0-5.5, tương đương trình độ trung cấp. Học viên có khả năng giao tiếp cơ bản và xử lý được các dạng bài thi quen thuộc, tuy nhiên cần rèn luyện thêm để xử lý tốt các dạng câu hỏi khó và nâng cao độ chính xác trong Nói và Viết.",
   "Học viên đã có nền tảng tương đối ổn định và đạt kết quả ở mức trung bình khá. Để tiến gần hơn tới mục tiêu Band điểm cao hơn, học viên cần luyện tập chuyên sâu các kỹ thuật làm bài và cải thiện độ trôi chảy, mạch lạc khi Nói và Viết.",
 ], "Học viên nên tập trung luyện các kỹ thuật làm bài (skimming, scanning, paraphrasing...) và tăng cường thực hành Nói - Viết theo chủ đề học thuật. Nên làm đề thi thử định kỳ để đánh giá tiến độ và điều chỉnh kế hoạch ôn tập kịp thời."),
 (6.5, "6.0–6.5", "Khá (trên trung cấp)", [
   "Học viên đạt Band 6.0-6.5, tương đương trình độ khá - trên trung cấp, đáp ứng tốt yêu cầu đầu vào của nhiều trường đại học quốc tế. Học viên xử lý bài thi khá thành thạo, tuy nhiên vẫn còn một số lỗi nhỏ về ngữ pháp hoặc cách triển khai ý cần hoàn thiện thêm.",
   "Học viên có nền tảng tiếng Anh khá vững và đạt kết quả tốt ở phần lớn các kỹ năng. Để đạt Band điểm cao hơn, học viên cần trau chuốt thêm về độ đa dạng từ vựng, cấu trúc câu và khả năng lập luận trong Nói - Viết.",
 ], "Học viên nên tiếp tục luyện tập nâng cao, tập trung vào việc đa dạng hóa từ vựng - cấu trúc câu và rèn luyện khả năng phản biện, lập luận trong Nói và Viết. Nên đặt mục tiêu làm quen với các dạng đề khó hơn để tiến gần Band điểm mục tiêu."),
 (999, "7.0+", "Cao cấp (xuất sắc)", [
   "Học viên đạt Band 7.0 trở lên, tương đương trình độ cao cấp, đáp ứng tốt yêu cầu của hầu hết các chương trình học thuật và du học quốc tế. Học viên thể hiện khả năng sử dụng tiếng Anh linh hoạt, tự nhiên và xử lý tốt các dạng đề khó.",
   "Đây là kết quả rất đáng khích lệ, cho thấy học viên đã có nền tảng tiếng Anh vững chắc và khả năng tư duy ngôn ngữ tốt. Học viên nên tiếp tục duy trì và trau dồi thêm để hướng đến Band điểm cao hơn nếu có nhu cầu.",
 ], "Học viên nên tiếp tục duy trì thói quen luyện tập, mở rộng vốn từ vựng chuyên sâu và rèn luyện thêm các kỹ thuật nâng cao nếu đặt mục tiêu Band điểm cao hơn (7.5+). Có thể tham gia các hoạt động học thuật bằng tiếng Anh (tranh biện, viết luận...) để phát triển toàn diện."),
]

def ielts_band_bucket(overall):
    for upper, key, label, items, kn in IELTS_BANDS:
        if overall < upper:
            return {"key": key, "label": label, "items": items, "kien_nghi": kn}
    k = IELTS_BANDS[-1]
    return {"key": k[1], "label": k[2], "items": k[3], "kien_nghi": k[4]}

# ---- Loi khuyen luyen tap theo tung ky nang (dung chung cho ca thang 0-10 va IELTS) ----
SKILL_TIPS = {
 "Nghe": "Nên luyện nghe thêm 2-3 giờ/tuần: nghe podcast/video tiếng Anh theo chủ đề quen thuộc, luyện nghe chép chính tả (dictation) đoạn ngắn, làm lại các dạng bài nghe trong đề thi để quen format và tốc độ nói.",
 "Nói": "Nên luyện nói thêm 2-3 giờ/tuần: luyện phát âm - ngữ điệu qua shadowing (nói theo audio mẫu), tập trả lời các câu hỏi theo chủ đề quen thuộc, ghi âm lại phần nói của mình để tự đánh giá và cải thiện.",
 "Đọc": "Nên luyện đọc thêm 2-3 giờ/tuần: đọc các đoạn văn ngắn phù hợp trình độ, luyện kỹ năng đọc lướt (skimming) và đọc quét (scanning) để tìm thông tin nhanh, mở rộng vốn từ vựng qua ngữ cảnh.",
 "Viết": "Nên luyện viết thêm 2-3 giờ/tuần: luyện viết câu, đoạn văn theo chủ đề quen thuộc, chú ý cấu trúc ngữ pháp và cách triển khai ý, tự soát lại bài và rút kinh nghiệm từ lỗi sai thường gặp.",
}

# ---- Thoi gian hoc toi uu goi y theo xep loai chuan (A+...D) ----
STUDY_TIME_STANDARD = {
 "A+": "2 buổi/tuần, 90 phút/buổi — duy trì nhịp học hiện tại.",
 "A":  "2 buổi/tuần, 90 phút/buổi — duy trì nhịp học hiện tại.",
 "B+": "2-3 buổi/tuần, 90 phút/buổi.",
 "B":  "2-3 buổi/tuần, 90 phút/buổi, nên thêm thời gian tự học ở nhà.",
 "C+": "3 buổi/tuần, 90-120 phút/buổi, kết hợp tự học thêm ở nhà.",
 "C":  "3 buổi/tuần, 90-120 phút/buổi, kết hợp tự học thêm ở nhà.",
 "D":  "3-4 buổi/tuần, 120 phút/buổi; nên cân nhắc thêm hỗ trợ kèm riêng (1-1) nếu có thể.",
}

def study_time_ielts(overall):
    if overall < 4.0:
        return "4 buổi/tuần, 120 phút/buổi, ưu tiên củng cố nền tảng trước khi luyện đề."
    if overall < 5.0:
        return "3-4 buổi/tuần, 120 phút/buổi."
    if overall < 6.0:
        return "3 buổi/tuần, 90-120 phút/buổi."
    if overall < 7.0:
        return "2-3 buổi/tuần, 90 phút/buổi."
    return "2 buổi/tuần, 90 phút/buổi — duy trì nhịp học hiện tại."

# ---- Bo nhan xet rieng cho LOP E (thieu nhi, chi hoc Noi + Chuyen can) ----
# Nguon: file chinh thuc "tiêu_chí_chấm_lớp_E.xlsx" - chia 3 muc Trung binh/Kha/Tot.
# Anh xa: Tot = A+/A, Kha = B+/B, Trung binh = C+/C, Yeu (tu soan them) = D.
_E_TOT = {"label": "Tốt", "items": [
   "Bé đi học đầy đủ, luôn tích cực tham gia các hoạt động và hợp tác tốt trong giờ học. Bé ghi nhớ tốt từ vựng, mẫu câu và có khả năng trả lời câu hỏi bằng tiếng Anh khá tự tin.",
   "Bé có chuyên cần tốt, ngoan và rất hào hứng trong các hoạt động trên lớp. Bé có khả năng phản xạ tốt, phát âm khá rõ và chủ động sử dụng tiếng Anh khi giao tiếp với cô.",
   "Bé có tinh thần học tập rất tốt, đi học đầy đủ và tích cực tham gia các hoạt động. Bé tự tin, chủ động và hợp tác tốt. Kỹ năng Speaking của bé rất tốt, bé ghi nhớ bài nhanh và có khả năng sử dụng tiếng Anh linh hoạt trong các tình huống quen thuộc.",
 ], "kien_nghi": [
   "Tiếp tục tạo cơ hội cho bé giao tiếp bằng tiếng Anh tại nhà để duy trì sự tự tin và phát triển khả năng phản xạ.",
   "Khuyến khích bé mở rộng câu trả lời thay vì chỉ sử dụng câu ngắn, giúp khả năng diễn đạt ngày càng tốt hơn.",
   "Có thể cho bé luyện hội thoại theo các chủ đề quen thuộc như gia đình, trường học, đồ chơi, con vật… để phát triển khả năng giao tiếp tự nhiên.",
   "Duy trì thói quen nghe và nói tiếng Anh mỗi ngày để bé phát huy tốt khả năng Speaking và sử dụng ngôn ngữ linh hoạt hơn.",
 ]}

_E_KHA = {"label": "Khá", "items": [
   "Bé đi học đầy đủ, ngoan và tích cực tham gia các hoạt động trên lớp. Bé nắm được các từ vựng và mẫu câu đã học, có thể trả lời các câu hỏi quen thuộc và đang dần tự tin hơn khi giao tiếp bằng tiếng Anh.",
   "Bé có chuyên cần tốt, tham gia hoạt động khá tích cực và hợp tác tốt với cô. Bé có khả năng ghi nhớ từ vựng, mẫu câu và trả lời được các câu hỏi cơ bản. Bé cần luyện thêm phát âm và phản xạ để giao tiếp tự nhiên hơn.",
   "Bé có thái độ học tập tốt, ngoan và phối hợp tốt trong các hoạt động trên lớp. Bé khá tự tin khi tham gia phần Speaking, biết sử dụng các mẫu câu đã học và có sự tiến bộ rõ rệt trong quá trình học.",
 ], "kien_nghi": [
   "Tiếp tục duy trì việc luyện nói tại nhà, đặc biệt là các mẫu câu đã học, để bé sử dụng tiếng Anh tự nhiên và linh hoạt hơn.",
   "Khuyến khích bé chủ động đặt câu hỏi và trả lời bằng tiếng Anh trong các tình huống quen thuộc.",
   "Ba mẹ có thể cho bé nghe các đoạn hội thoại ngắn và khuyến khích bé lặp lại để cải thiện phát âm và ngữ điệu.",
   "Nên duy trì thói quen luyện Speaking hằng ngày từ 5–10 phút để bé củng cố kiến thức và tăng phản xạ.",
 ]}

_E_TB = {"label": "Trung bình", "items": [
   "Bé đi học khá đầy đủ, ngoan và có tham gia các hoạt động trên lớp. Bé đôi lúc còn rụt rè khi sử dụng tiếng Anh và cần cô hỗ trợ thêm khi trả lời câu hỏi. Bé cần luyện tập thêm từ vựng và mẫu câu để tự tin hơn khi giao tiếp.",
   "Bé có ý thức đi học và tham gia các hoạt động cùng cô và các bạn. Bé hợp tác khá tốt nhưng đôi khi còn mất tập trung. Kỹ năng Speaking của bé ở mức khá cơ bản, cần luyện tập thêm để tăng khả năng phản xạ và phát âm.",
   "Bé đi học tương đối đầy đủ, ngoan và biết hợp tác trong giờ học. Bé đã ghi nhớ được một số từ vựng và mẫu câu quen thuộc nhưng còn cần nhắc khi trả lời. Cô mong bé mạnh dạn nói tiếng Anh nhiều hơn trong thời gian tới.",
 ], "kien_nghi": [
   "Khuyến khích bé luyện nói tiếng Anh thường xuyên tại nhà thông qua việc nghe và lặp lại các câu mẫu đã học trên lớp.",
   "Ba mẹ có thể dành thời gian ôn lại từ vựng và mẫu câu cùng bé mỗi ngày để bé ghi nhớ lâu hơn.",
   "Khuyến khích bé trả lời bằng tiếng Anh thay vì chỉ trả lời ngắn hoặc bằng tiếng Việt, giúp bé hình thành phản xạ giao tiếp tự nhiên.",
   "Bé cần được tạo môi trường giao tiếp tiếng Anh thường xuyên hơn để tăng sự tự tin và khả năng phản xạ khi nói.",
 ]}

_E_YEU = {"label": "Yếu", "items": [
   "Bé còn khá rụt rè và ít tham gia các hoạt động nói trên lớp, cần được cô hỗ trợ và khích lệ nhiều hơn để mạnh dạn sử dụng tiếng Anh.",
   "Bé còn gặp khó khăn trong việc ghi nhớ từ vựng và mẫu câu, khả năng phản xạ khi giao tiếp còn hạn chế. Bé cần thêm thời gian và sự đồng hành sát sao từ cô và gia đình.",
   "Bé đi học chưa đều hoặc chưa thật sự tập trung trong giờ học, ảnh hưởng đến khả năng tiếp thu từ vựng và mẫu câu. Cần có kế hoạch hỗ trợ riêng để bé bắt kịp các bạn.",
 ], "kien_nghi": [
   "Ba mẹ nên dành thời gian luyện nói cùng bé mỗi ngày, bắt đầu từ những từ vựng và mẫu câu đơn giản, quen thuộc.",
   "Khuyến khích bé nghe tiếng Anh qua bài hát, video ngắn phù hợp lứa tuổi để làm quen phát âm một cách tự nhiên, không áp lực.",
   "Nên kiên nhẫn động viên bé, tránh tạo áp lực khi bé chưa nói được, giúp bé cảm thấy thoải mái và dần tự tin hơn.",
   "Ba mẹ có thể cân nhắc thêm thời gian luyện tập ngoài giờ học nếu bé cần hỗ trợ nhiều hơn để bắt kịp các bạn.",
 ]}

E_CLASS = {
 "A+": _E_TOT,
 "A": _E_TOT,
 "B+": _E_KHA,
 "B": _E_KHA,
 "C+": _E_TB,
 "C": _E_TB,
 "D": _E_YEU,
}

import re

SKILL_PAT = re.compile(r'^(.*?)\n.*?\((\d+(?:\.\d+)?)%\)\s*$', re.S)
MAX_PAT = re.compile(r'\((\d+(?:\.\d+)?)\)')
CC_PAT = re.compile(r'Chuyên cần.*?\((\d+(?:\.\d+)?)%\)', re.S)

def find_header_row(ws, max_scan=40):
    for r in range(1, min(max_scan, ws.max_row) + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if v and str(v).strip().startswith('STT'):
                return r
    return None

def col_text(ws, r, c):
    v = ws.cell(r, c).value
    return str(v) if v is not None else ''

def is_cc_sheet(sheet_name):
    return bool(re.match(r'^\s*\d+\.\s*CC\s', sheet_name))

def detect_sheet_kind(ws, hdr_row):
    row_text = ' '.join(col_text(ws, hdr_row, c) for c in range(1, ws.max_column + 1))
    row_text += ' ' + ' '.join(col_text(ws, hdr_row + 1, c) for c in range(1, ws.max_column + 1))
    if 'OVERALL BAND SCORE' in row_text.upper() or 'Band score' in row_text:
        return 'ielts'
    return 'standard'

def parse_cc_sheet(ws, hdr_row):
    # header spans hdr_row, hdr_row+1 ; data starts hdr_row+2 typically (numbers row) -> hdr_row+3
    ncols = ws.max_column
    layout = {'stt': None, 'ten': None, 'tong_col': None,
              'kyluat_col': None, 'diemdanh_col': None, 'btvn_col': None, 'hoatdong_col': None}
    for c in range(1, ncols + 1):
        t = col_text(ws, hdr_row, c)
        if t.strip().startswith('STT'):
            layout['stt'] = c
        if 'học viên' in t.lower() and 'tên' in t.lower():
            layout['ten'] = c
        t2 = col_text(ws, hdr_row + 1, c)
        if 'Tổng cộng' in t2:
            layout['tong_col'] = c
        if 'Kỷ luật' in t2:
            layout['kyluat_col'] = c
        if 'Điểm danh' in t2:
            layout['diemdanh_col'] = c
        if 'Bài tập về nhà' in t2 or 'BTVN' in t2:
            layout['btvn_col'] = c
        if 'Hoạt động' in t2:
            layout['hoatdong_col'] = c
    data_start = hdr_row + 3  # skip the "1,2,3.." index row
    return layout, data_start

def parse_standard_sheet(ws, hdr_row):
    r1, r2, r3 = hdr_row, hdr_row + 1, hdr_row + 2
    ncols = ws.max_column
    layout = {'stt': None, 'ten': None, 'cc_col': None, 'cc_weight': None,
              'skills': [], 'tong_col': None, 'kyhieu_col': None, 'xeploai_col': None, 'ghichu_col': None}
    for c in range(1, ncols + 1):
        t1 = col_text(ws, r1, c)
        if t1.strip().startswith('STT'):
            layout['stt'] = c
        if t1.strip().startswith('Tên'):
            layout['ten'] = c
        m = CC_PAT.search(t1)
        if m:
            layout['cc_col'] = c
            layout['cc_weight'] = float(m.group(1))
        if 'Tổng điểm' in t1 or 'Total score' in t1:
            layout['tong_col'] = c
        if 'Ghi chú' in t1:
            layout['ghichu_col'] = c
    for c in range(1, ncols + 1):
        t2 = col_text(ws, r2, c)
        m = SKILL_PAT.match(t2)
        if m:
            name_vn = m.group(1).strip()
            weight = float(m.group(2))
            t3 = col_text(ws, r3, c)
            mm = MAX_PAT.search(t3)
            max_score = float(mm.group(1)) if mm else None
            layout['skills'].append({'name': name_vn, 'weight': weight, 'max': max_score, 'col': c})
    for c in range(1, ncols + 1):
        t3 = col_text(ws, r3, c)
        if 'Ký hiệu' in t3:
            layout['kyhieu_col'] = c
        if t3.strip() == 'Xếp loại':
            layout['xeploai_col'] = c
    return layout, r3 + 1

def parse_ielts_sheet(ws, hdr_row):
    r1, r2, r3 = hdr_row, hdr_row + 1, hdr_row + 2
    ncols = ws.max_column
    layout = {'stt': None, 'ten': None, 'skills': [], 'overall_col': None, 'ghichu_col': None}
    for c in range(1, ncols + 1):
        t1 = col_text(ws, r1, c)
        if t1.strip().startswith('STT'):
            layout['stt'] = c
        if t1.strip().startswith('Tên'):
            layout['ten'] = c
        if 'OVERALL BAND' in t1.upper():
            layout['overall_col'] = c
        if 'Ghi chú' in t1:
            layout['ghichu_col'] = c
    for c in range(1, ncols + 1):
        t3 = col_text(ws, r3, c).strip()
        if t3 == 'Band score':
            # find nearest skill label scanning left in row r2
            name = None
            for back in range(0, 5):
                cc = c - back
                if cc < 1:
                    break
                t2 = col_text(ws, r2, cc).strip()
                if t2:
                    name = t2.split('\n')[0].strip()
                    break
            layout['skills'].append({'name': name, 'band_col': c})
    return layout, r3 + 2  # skip the extra "1,2,3.." position-index filler row

def _num(v):
    try:
        if v is None or v == '':
            return None
        return float(v)
    except (TypeError, ValueError):
        return None

def build_cc_lookup(wb_v, cc_sheet_name):
    ws = wb_v[cc_sheet_name]
    hdr = find_header_row(ws)
    layout, data_start = parse_cc_sheet(ws, hdr)
    lut = {}
    r = data_start
    blanks = 0
    while r <= ws.max_row and blanks < 3:
        name = ws.cell(r, layout['ten']).value if layout['ten'] else None
        if name is None or str(name).strip() == '':
            blanks += 1
            r += 1
            continue
        blanks = 0
        tong = _num(ws.cell(r, layout['tong_col']).value) if layout['tong_col'] else None
        entry = {'tong': tong}
        for key in ('kyluat_col', 'diemdanh_col', 'btvn_col', 'hoatdong_col'):
            col = layout.get(key)
            entry[key.replace('_col', '')] = _num(ws.cell(r, col).value) if col else None
        lut[str(name).strip().upper()] = entry
        r += 1
    return lut

def extract_standard(wb_f, wb_v, sheet_name, class_info, cc_lookup=None):
    ws_f = wb_f[sheet_name]
    ws_v = wb_v[sheet_name]
    hdr = find_header_row(ws_f)
    layout, data_start = parse_standard_sheet(ws_f, hdr)
    students = []
    r = data_start
    blanks = 0
    while r <= ws_v.max_row and blanks < 3:
        name = ws_v.cell(r, layout['ten']).value if layout['ten'] else None
        stt_val = _num(ws_v.cell(r, layout['stt']).value) if layout['stt'] else None
        if name is None or str(name).strip() == '' or stt_val is None:
            blanks += 1
            r += 1
            continue
        blanks = 0
        name = str(name).strip()
        cc_val = _num(ws_v.cell(r, layout['cc_col']).value) if layout['cc_col'] else None
        cc_breakdown = None
        if cc_lookup is not None:
            entry = cc_lookup.get(name.upper())
            if entry:
                cc_breakdown = entry
                if cc_val is None:
                    cc_val = entry.get('tong')
        if cc_val is None:
            cc_val = 0.0

        # Neu sheet Chuyen can khong co / khong co du lieu chi tiet (Ky luat, Diem danh, BTVN, Hoat dong
        # deu rong) -> "tinh nguoc" bang cach dung diem Chuyen can tong o file chinh cho ca 4 muc con,
        # de bang 1 tren phieu luon co du liệu hien thi thay vi de trong.
        breakdown_missing = (cc_breakdown is None) or all(
            cc_breakdown.get(k) is None for k in ('kyluat', 'diemdanh', 'btvn', 'hoatdong')
        )
        if breakdown_missing:
            cc_breakdown = {'kyluat': cc_val, 'diemdanh': cc_val, 'btvn': cc_val, 'hoatdong': cc_val, 'tong': cc_val}

        skills = []
        for sk in layout['skills']:
            raw = _num(ws_v.cell(r, sk['col']).value)
            maxv = sk['max'] or 10.0
            if raw is None:
                # Khong co diem = None, khong phai 0.
                skills.append({'name': sk['name'], 'raw': None, 'max': maxv, 'weight': sk['weight'],
                                'pct_contrib': None, 'norm10': None})
            else:
                pct_contrib = round((raw / maxv) * sk['weight'], 1) if maxv else 0.0
                normalized10 = round((raw / maxv) * 10, 2) if maxv else 0.0
                skills.append({'name': sk['name'], 'raw': raw, 'max': maxv, 'weight': sk['weight'],
                                'pct_contrib': pct_contrib, 'norm10': normalized10})

        # Neu thieu diem o bat ky ky nang nao: KHONG tinh lai tong/xep loai.
        # Chi danh dau hoc vien chua du du lieu de tinh ket qua cuoi khoa.
        available = [s for s in skills if s['raw'] is not None]
        missing_skills = [s['name'] for s in skills if s['raw'] is None]
        if missing_skills:
            total = None
            grade = 'Chưa đủ dữ liệu điểm số, yêu cầu cập nhật lại điểm số'
        else:
            cached_tong = _num(ws_v.cell(r, layout['tong_col']).value) if layout['tong_col'] else None
            test_contrib = sum(s['pct_contrib'] for s in available)
            if cached_tong is not None:
                total = cached_tong
            else:
                total = round(test_contrib + cc_val, 1)
            cached_grade = ws_v.cell(r, layout['kyhieu_col']).value if layout['kyhieu_col'] else None
            grade = cached_grade.strip() if isinstance(cached_grade, str) and cached_grade.strip() else grade_from_total(total)

        # Chi danh gia ky nang co diem. Ky nang 'Chua co diem' khong duoc xem la ky nang yeu.
        weak = []
        if len(available) >= 2:
            norms = [s['norm10'] for s in available]
            for s in available:
                others_avg = (sum(norms) - s['norm10']) / (len(norms) - 1)
                if s['norm10'] < 6 and s['norm10'] <= others_avg - 1:
                    weak.append(s['name'])

        students.append({
            'name': name, 'type': 'standard', 'class_info': class_info,
            'cc': cc_val, 'cc_weight': layout.get('cc_weight') or 10.0, 'cc_breakdown': cc_breakdown,
            'skills': skills, 'total': total, 'grade': grade,
            'weak_skills': weak,
        })
        r += 1
    return students

def extract_ielts(wb_f, wb_v, sheet_name, class_info):
    ws_f = wb_f[sheet_name]
    ws_v = wb_v[sheet_name]
    hdr = find_header_row(ws_f)
    layout, data_start = parse_ielts_sheet(ws_f, hdr)
    students = []
    r = data_start
    blanks = 0
    while r <= ws_v.max_row and blanks < 3:
        name = ws_v.cell(r, layout['ten']).value if layout['ten'] else None
        stt_val = _num(ws_v.cell(r, layout['stt']).value) if layout['stt'] else None
        if name is None or str(name).strip() == '' or stt_val is None:
            blanks += 1
            r += 1
            continue
        blanks = 0
        name = str(name).strip()
        skills = []
        for sk in layout['skills']:
            band = _num(ws_v.cell(r, sk['band_col']).value)
            skills.append({'name': sk['name'], 'band': band})
        available = [s for s in skills if s['band'] is not None]
        missing_skills = [s['name'] for s in skills if s['band'] is None]
        cached_overall = _num(ws_v.cell(r, layout['overall_col']).value) if layout['overall_col'] else None
        # Neu thieu bat ky ky nang nao: khong tinh lai Overall.
        # Ket qua se thong bao chua du du lieu va yeu cau cap nhat diem.
        if missing_skills:
            overall = None
        elif cached_overall is not None:
            overall = cached_overall
        elif available:
            overall = round(sum(s['band'] for s in available) / len(available), 1)
        else:
            overall = None

        # Chi danh gia ky nang co diem.
        weak = []
        if len(available) >= 2:
            bands = [s['band'] for s in available]
            for s in available:
                others_avg = (sum(bands) - s['band']) / (len(bands) - 1)
                if s['band'] < 5.0 and s['band'] <= others_avg - 1.0:
                    weak.append(s['name'])

        students.append({
            'name': name, 'type': 'ielts', 'class_info': class_info,
            'skills': skills, 'overall': overall, 'weak_skills': weak,
        })
        r += 1
    return students

def process_workbook(path):
    """Returns list of student dicts across all class sheets found in the workbook."""
    wb_f = openpyxl.load_workbook(path, data_only=False)
    wb_v = openpyxl.load_workbook(path, data_only=True)
    all_students = []
    sheet_names = wb_f.sheetnames
    for name in sheet_names:
        if is_cc_sheet(name):
            continue
        ws = wb_f[name]
        hdr = find_header_row(ws)
        if hdr is None:
            continue
        # class info from rows above header (Lop/Giao vien/Khoa hoc)
        class_info = {'sheet': name, 'lop': None, 'giaovien': None, 'ngay_thi': None}
        for rr in range(1, hdr):
            v = ws.cell(rr, 1).value
            if v and isinstance(v, str):
                if v.strip().startswith('Lớp'):
                    class_info['lop'] = v.split(':', 1)[-1].strip()
                if v.strip().startswith('Giáo viên'):
                    class_info['giaovien'] = v.split(':', 1)[-1].strip()
                if v.strip().startswith('Ngày thi'):
                    class_info['ngay_thi'] = v.split(':', 1)[-1].strip()
        kind = detect_sheet_kind(ws, hdr)
        if kind == 'standard':
            cc_sheet_guess = re.sub(r'^(\s*\d+\.\s*)', r'\1CC ', name)
            cc_lookup = None
            if cc_sheet_guess in sheet_names:
                cc_lookup = build_cc_lookup(wb_v, cc_sheet_guess)
            all_students.extend(extract_standard(wb_f, wb_v, name, class_info, cc_lookup))
        else:
            all_students.extend(extract_ielts(wb_f, wb_v, name, class_info))
    return all_students

BRAND = RGBColor(0x1F, 0x4E, 0x79)
RED = RGBColor(0xC0, 0x00, 0x00)
COMPANY_LOGO_PATH = None  # set this (a file path) before calling build_phieu to insert a logo
LEGEND_LEFT = ["A+   Outstanding (Xuất sắc): 98 – 100", "A     Excellent (Rất giỏi): 95 - 97", "B+   Very good (Giỏi): 90 – 94"]
LEGEND_RIGHT = ["B     Good (Khá): 80 - 89", "C+   Average (Trung bình): 75 - 79", "C     Need improvement (Cần cố gắng hơn): 65 - 74", "D     Fail (Không đạt) < 65"]


def _shade(cell, hex_color):
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), hex_color)
    cell._tc.get_or_add_tcPr().append(shd)


def _vcenter(cell):
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


def _set_cell_text(cell, text, bold=False, italic=False, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT, color=None):
    cell.text = ''
    p = cell.paragraphs[0]
    p.alignment = align
    lines = str(text).split('\n')
    for i, line in enumerate(lines):
        if i > 0:
            p.add_run().add_break()
        r = p.add_run(line)
        r.bold = bold
        r.italic = italic
        r.font.size = Pt(size)
        if color:
            r.font.color.rgb = color
    _vcenter(cell)
    return cell


def _merge_col_widths(table, widths_cm):
    for row in table.rows:
        for idx, w in enumerate(widths_cm):
            row.cells[idx].width = Cm(w)


def _section_title_row(table, row_idx, text, ncols, size=11):
    row = table.rows[row_idx]
    a = row.cells[0]
    for i in range(1, ncols):
        a = a.merge(row.cells[i])
    _set_cell_text(a, text, bold=True, size=size, align=WD_ALIGN_PARAGRAPH.LEFT)
    _shade(a, 'D9E2F3')


def _add_logo_header(doc):
    """Chèn logo vào header của document -> tự động lặp lại trên MỌI trang,
    kể cả các trang được thêm sau này (Nhận xét, v.v.), và không bị thay thế khi nội dung thay đổi."""
    if not COMPANY_LOGO_PATH:
        return
    section = doc.sections[0]
    section.header.is_linked_to_previous = False
    # xoa noi dung header cu (neu co) va chen logo
    header_p = section.header.paragraphs[0]
    header_p.text = ''
    header_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = header_p.add_run()
    try:
        run.add_picture(COMPANY_LOGO_PATH, width=Cm(2.6))
    except Exception:
        pass


def _add_logo_and_title(doc):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("BÁO CÁO KẾT QUẢ HỌC TẬP CUỐI KHÓA")
    r.bold = True
    r.font.size = Pt(16)
    p.paragraph_format.space_after = Pt(0)

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run("(FINAL COURSE REPORT)")
    r2.bold = True
    r2.italic = True
    r2.font.size = Pt(13)
    p2.paragraph_format.space_after = Pt(10)


def _add_student_info(doc, student):
    p = doc.add_paragraph()
    r1 = p.add_run("Tên học viên (Student's name): ")
    r1.font.size = Pt(11)
    r2 = p.add_run(student['name'])
    r2.bold = True
    r2.font.size = Pt(12)
    p.paragraph_format.space_after = Pt(4)

    table = doc.add_table(rows=1, cols=4)
    table.autofit = True
    cells = table.rows[0].cells
    _set_cell_text(cells[0], "Lớp (Class):", bold=False, size=10.5)
    _set_cell_text(cells[1], student['class_info'].get('lop') or '', bold=True, size=11)
    _set_cell_text(cells[2], "Ngày thi (Date):", bold=False, size=10.5)
    _set_cell_text(cells[3], student['class_info'].get('ngay_thi') or '', bold=True, size=11)
    for row in table.rows:
        for c in row.cells:
            c._tc.get_or_add_tcPr()
    table.style = 'Table Grid'
    for row in table.rows:
        for c in row.cells:
            tcPr = c._tc.get_or_add_tcPr()
            borders = OxmlElement('w:tcBorders')
            for edge in ('top', 'left', 'bottom', 'right'):
                el = OxmlElement(f'w:{edge}')
                el.set(qn('w:val'), 'nil')
                borders.append(el)
            tcPr.append(borders)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def _add_table1_attendance(doc, student):
    bd = student.get('cc_breakdown')
    if bd:
        rows_data = [
            ("Kỷ luật\n(Discipline)", bd.get('kyluat')),
            ("Điểm danh\n(Attendance)", bd.get('diemdanh')),
            ("Bài tập về nhà\n(Homework)", bd.get('btvn')),
            ("Hoạt động trong lớp\n(Activity)", bd.get('hoatdong')),
        ]
    else:
        rows_data = [("Chuyên cần\n(Participation)", student['cc'])]

    n = len(rows_data)
    table = doc.add_table(rows=1 + n, cols=4)
    table.style = 'Table Grid'
    _section_title_row(table, 0, "1. ĐÁNH GIÁ CỦA GIÁO VIÊN TRONG QUÁ TRÌNH HỌC (Teacher's evaluation) (10%)", 4, size=10.5)

    left = table.cell(1, 0)
    for i in range(2, 1 + n):
        left = left.merge(table.cell(i, 0))
    _set_cell_text(left, "Chuyên cần\n(Class participation)", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER)

    right = table.cell(1, 3)
    for i in range(2, 1 + n):
        right = right.merge(table.cell(i, 3))
    cc_weight = student.get('cc_weight') or 10.0
    contrib = round((student['cc'] / 10.0) * cc_weight, 1)
    contrib_txt = f"{contrib:.0f}%" if contrib == int(contrib) else f"{contrib:.1f}%"
    avg_txt = f"Điểm trung bình chuyên cần\n(Average)\n\n{contrib_txt}"
    _set_cell_text(right, avg_txt, bold=False, size=10, align=WD_ALIGN_PARAGRAPH.CENTER)

    for i, (label, val) in enumerate(rows_data):
        r = 1 + i
        _set_cell_text(table.cell(r, 1), label, size=10, align=WD_ALIGN_PARAGRAPH.LEFT)
        vtxt = f"{val:.1f}/10" if val is not None else "-"
        _set_cell_text(table.cell(r, 2), vtxt, bold=False, size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def _add_table2_test_result(doc, student):
    skills = student['skills']
    n = len(skills)
    table = doc.add_table(rows=2 + n, cols=3)
    table.style = 'Table Grid'
    _section_title_row(table, 0, "2. KẾT QUẢ BÀI THI (Test Result) (90%)", 3, size=10.5)
    _set_cell_text(table.cell(1, 0), "", size=10)
    _set_cell_text(table.cell(1, 1), "Điểm thi\n(Test Score)", bold=True, size=10, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text(table.cell(1, 2), "Kết quả\n(Result in Percentage)", bold=True, size=10, align=WD_ALIGN_PARAGRAPH.CENTER)
    for i, sk in enumerate(skills):
        r = 2 + i
        eng = {"Nói": "Speaking", "Nghe": "Listening", "Đọc": "Reading", "Viết": "Writing"}.get(sk['name'], sk['name'])
        _set_cell_text(table.cell(r, 0), f"{sk['name']} ({eng}): {sk['weight']:.0f}%", bold=True, size=10.5)
        if sk['raw'] is None:
            _set_cell_text(table.cell(r, 1), "Chưa có điểm", italic=True, size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER)
            _set_cell_text(table.cell(r, 2), "-", size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER)
        else:
            _set_cell_text(table.cell(r, 1), f"{sk['raw']:.0f}/{sk['max']:.0f}", size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER)
            _set_cell_text(table.cell(r, 2), f"{sk['pct_contrib']:.1f}%", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def _add_table3_final(doc, student):
    table = doc.add_table(rows=6, cols=2)
    table.style = 'Table Grid'
    _section_title_row(table, 0, "3. KẾT QUẢ CUỐI KHÓA (Final result) (100%)", 2, size=10.5)
    _set_cell_text(table.cell(1, 0), "Tổng cộng (Total):", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER)
    if student.get('total') is None:
        total_txt = "Chưa đủ dữ liệu điểm số, yêu cầu cập nhật lại điểm số"
        total_size = 10
    else:
        total_txt = f"{student['total']:.1f}".replace('.', ',')
        total_size = 12
    _set_cell_text(table.cell(1, 1), total_txt, bold=True, size=total_size, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text(table.cell(2, 0), "Xếp loại (Rating):", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text(table.cell(2, 1), student['grade'], bold=True, size=13, align=WD_ALIGN_PARAGRAPH.CENTER, color=RED)

    c0 = table.cell(3, 0)
    _set_cell_text(c0, "Ghi chú (Note):\n" + "\n".join(LEGEND_LEFT), size=8.5, italic=True)
    c1 = table.cell(3, 1)
    _set_cell_text(c1, "\n" + "\n".join(LEGEND_RIGHT), size=8.5, italic=True)

    nx = table.cell(4, 0)
    nx = nx.merge(table.cell(4, 1))
    _set_cell_text(nx, "Nhận xét khác (Other comments)\n\n\n", size=10, align=WD_ALIGN_PARAGRAPH.CENTER)

    _set_cell_text(table.cell(5, 0), "Chữ ký GVCN\n(Teacher's signature & full name)", size=9.5, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text(table.cell(5, 1), student['class_info'].get('giaovien') or '', bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def _bao_cao_date(student):
    """Ngay ghi o cuoi phieu = ngay ket thuc thi (lay tu 'Ngày thi' trong file diem) + 2 ngay.
    Neu khong doc duoc ngay thi, dung ngay hom nay lam du phong."""
    ngay_thi_str = student.get('class_info', {}).get('ngay_thi') or ''
    matches = re.findall(r'(\d{1,2})/(\d{1,2})/(\d{4})', ngay_thi_str)
    if matches:
        d, m, y = matches[-1]  # ngay cuoi cung trong chuoi la ngay ket thuc thi
        try:
            end_date = datetime.date(int(y), int(m), int(d))
            return end_date + datetime.timedelta(days=2)
        except ValueError:
            pass
    return datetime.date.today()


def _add_footer_signatures(doc, student):
    report_date = _bao_cao_date(student)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p.add_run(f"Ngày {report_date.day} tháng {report_date.month} năm {report_date.year}")
    r.font.size = Pt(10.5)

    table = doc.add_table(rows=3, cols=3)
    for row in table.rows:
        for c in row.cells:
            tcPr = c._tc.get_or_add_tcPr()
            borders = OxmlElement('w:tcBorders')
            for edge in ('top', 'left', 'bottom', 'right'):
                el = OxmlElement(f'w:{edge}')
                el.set(qn('w:val'), 'nil')
                borders.append(el)
            tcPr.append(borders)
    _set_cell_text(table.cell(0, 0), "Trưởng bộ phận học vụ", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text(table.cell(0, 1), "Người lập", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text(table.cell(0, 2), "GVCN", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text(table.cell(1, 0), "(Chief Academic Officer)", italic=True, size=9.5, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text(table.cell(1, 1), "(Prepared by)", italic=True, size=9.5, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text(table.cell(1, 2), "(Teacher's signature & full name)", italic=True, size=9.5, align=WD_ALIGN_PARAGRAPH.CENTER)
    # Chi de ten GVCN o cot GVCN; khong dien ten o Truong bo phan hoc vu va Nguoi lap.
    gv = student['class_info'].get('giaovien') or ''
    _set_cell_text(table.cell(2, 0), "\n", size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text(table.cell(2, 1), "\n", size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
    _set_cell_text(table.cell(2, 2), "\n" + gv, bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER)


# ---------------- Phan Nhan xet AI (giu nguyen phong cach cu, dat ben duoi bao cao chinh) ----------------

def _add_heading(doc, text, size=11.5, color=None, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, space_after=4):
    p = doc.add_paragraph()
    p.alignment = align
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.font.bold = bold
    if color:
        r.font.color.rgb = color
    p.paragraph_format.space_after = Pt(space_after)
    return p


def _add_section_bar(doc, text):
    """Thanh tieu de nen xanh nhat, dong bo voi cac muc 1./2./3. o trang 1."""
    table = doc.add_table(rows=1, cols=1)
    table.style = 'Table Grid'
    _set_cell_text(table.cell(0, 0), text, bold=True, size=11, align=WD_ALIGN_PARAGRAPH.LEFT)
    _shade(table.cell(0, 0), 'D9E2F3')
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def _add_body(doc, text, size=11, space_after=8):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.font.size = Pt(size)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.15
    return p


def _add_ai_section(doc, student):
    doc.add_page_break()
    _add_section_bar(doc, "NHẬN XÉT & LỘ TRÌNH ÔN TẬP")

    # Neu chua du diem, khong xep loai va khong tao nhan xet theo xep loai.
    if (student['type'] == 'standard' and student.get('total') is None) or (student['type'] == 'ielts' and student.get('overall') is None):
        _add_heading(doc, "Thông báo", size=11.5, space_after=2)
        _add_body(doc, "Chưa đủ dữ liệu điểm số, yêu cầu cập nhật lại điểm số.", size=10.5)
        return

    if student['type'] == 'standard':
        if len(student['skills']) == 1 and student['skills'][0]['name'].strip() == 'Nói':
            bucket = E_CLASS[student['grade']]
        else:
            bucket = STANDARD[student['grade']]
    else:
        bucket = ielts_band_bucket(student['overall'])

    _add_heading(doc, "I. Nhận xét", size=11.5, space_after=2)
    _add_body(doc, random.choice(bucket['items']), size=10.5)

    _add_heading(doc, "II. Kiến nghị", size=11.5, space_after=2)
    kn = bucket['kien_nghi']
    if isinstance(kn, list):
        for item in kn:
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.5)
            r = p.add_run(f"• {item}")
            r.font.size = Pt(10.5)
            p.paragraph_format.space_after = Pt(4)
    else:
        _add_body(doc, kn, size=10.5)

    _add_heading(doc, "III. Lộ trình ôn tập đề xuất", size=11.5, space_after=2)
    if student['weak_skills']:
        _add_body(doc, "Các kỹ năng cần ưu tiên cải thiện: " + ", ".join(student['weak_skills']) + ".", size=10.5, space_after=3)
        for sk_name in student['weak_skills']:
            tip = SKILL_TIPS.get(sk_name.strip())
            if tip:
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Cm(0.5)
                r = p.add_run(f"• {sk_name.strip()}: ")
                r.bold = True
                r.font.size = Pt(10.5)
                r2 = p.add_run(tip)
                r2.font.size = Pt(10.5)
                p.paragraph_format.space_after = Pt(4)
    else:
        if student['type'] == 'standard' and len(student['skills']) == 1:
            _add_body(doc, f"Học viên nên tiếp tục luyện tập kỹ năng {student['skills'][0]['name']} thường xuyên để duy trì và nâng cao phong độ hiện tại.", size=10.5, space_after=4)
        else:
            _add_body(doc, "Các kỹ năng của học viên hiện tương đối đồng đều, chưa ghi nhận kỹ năng nào cần ưu tiên đặc biệt. Học viên nên tiếp tục luyện tập đều cả các kỹ năng để duy trì phong độ.", size=10.5, space_after=4)

    _add_heading(doc, "IV. Thời gian học tập tối ưu đề xuất", size=11.5, space_after=2)
    if student['type'] == 'standard':
        _add_body(doc, STUDY_TIME_STANDARD.get(student['grade'], "2-3 buổi/tuần, 90 phút/buổi."), size=10.5)
    else:
        _add_body(doc, study_time_ielts(student['overall']), size=10.5)


def build_phieu(student):
    doc = Document()
    section = doc.sections[0]
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)
    section.top_margin = Cm(1.3)
    section.bottom_margin = Cm(1.3)

    for st in doc.styles:
        try:
            st.font.name = 'Times New Roman'
        except Exception:
            pass

    _add_logo_header(doc)
    _add_logo_and_title(doc)
    _add_student_info(doc, student)

    if student['type'] == 'standard':
        _add_table1_attendance(doc, student)
        _add_table2_test_result(doc, student)
        _add_table3_final(doc, student)
        _add_footer_signatures(doc, student)
    else:
        # IELTS: simpler score table (khong theo mau A-D), giu logo + tieu de dong bo
        table = doc.add_table(rows=1 + len(student['skills']), cols=2)
        table.style = 'Table Grid'
        _section_title_row(table, 0, "KẾT QUẢ BÀI THI IELTS (Band Score)", 2, size=10.5)
        for i, sk in enumerate(student['skills']):
            r = 1 + i
            _set_cell_text(table.cell(r, 0), sk['name'], bold=True, size=10.5)
            if sk['band'] is None:
                _set_cell_text(table.cell(r, 1), "Chưa có điểm", italic=True, size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER)
            else:
                _set_cell_text(table.cell(r, 1), f"{sk['band']:.1f}", bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
        doc.add_paragraph()
        _add_label = doc.add_paragraph()
        r1 = _add_label.add_run("Overall Band Score: ")
        r1.bold = True
        if student.get('overall') is None:
            r2 = _add_label.add_run("Chưa đủ dữ liệu điểm số, yêu cầu cập nhật lại điểm số")
            r2.bold = True
            r2.font.size = Pt(11)
        else:
            r2 = _add_label.add_run(f"{student['overall']:.1f}")
            r2.bold = True
            r2.font.size = Pt(13)
            r2.font.color.rgb = RED
        _add_footer_signatures(doc, student)

    _add_ai_section(doc, student)
    return doc

# ---- Anh nen he thong (CO DINH cho tat ca moi nguoi) ----
# De trong cho toi khi co file anh - dan chuoi base64 cua anh vao giua 2 dau """ ben duoi.
DEFAULT_BG_B64 = """"""

# ---- Logo mac dinh (CIE VIETNAM), nhung base64 de khong can file rieng khi deploy ----
DEFAULT_LOGO_B64 = """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"""
